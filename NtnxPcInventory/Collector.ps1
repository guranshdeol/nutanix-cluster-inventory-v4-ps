<#
    All Prism Central v4 API calls and data shaping.

    Column names and derived values mirror the Python collector so the workbook
    is interchangeable between the two implementations.
#>

Set-StrictMode -Version Latest

# Stats window settings - keep in step with the Python collector.
$script:SamplingInterval = 3600 * 4     # 4h buckets over a 24h window
$script:StatType         = 'AVG'
$script:VmStatsWorkers   = 3            # PC rate-limits hard; more workers is slower

$script:ColorUnderProvisioned = 'RED'
$script:ColorOverProvisioned  = 'YELLOW'
$script:ColorAdequate         = 'NONE'

$script:ClusterEnvKey  = 'cluster_environment'
$script:LocationKey    = 'location'
$script:DatacenterKey  = 'datacenter'
$script:RegionKey      = 'region'

# --- Categories -------------------------------------------------------------------

function Resolve-NtnxCategory {
    <# Returns "key:value" for a category extId, memoised per PC. #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$ExtId
    )

    if ($Session.CategoryCache.ContainsKey($ExtId)) { return $Session.CategoryCache[$ExtId] }

    $result = $null
    $path = Get-NtnxPath -Session $Session -Namespace 'prism' -Suffix "config/categories/$ExtId"
    if ($path) {
        try {
            $resp = Invoke-NtnxApi -Session $Session -Path $path -Quiet
            $key = Get-Prop $resp 'data.key' $null
            $value = Get-Prop $resp 'data.value' $null
            if ($null -ne $key) { $result = "${key}:${value}" }
        }
        catch {
            $result = $null
        }
    }

    $Session.CategoryCache[$ExtId] = $result
    return $result
}

# --- Clusters ---------------------------------------------------------------------

function Get-NtnxClusters {
    param([Parameter(Mandatory)]$Session)

    $path = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix 'config/clusters'
    if (-not $path) { throw "clustermgmt namespace is not available on $($Session.Host)." }
    return Get-NtnxList -Session $Session -Path $path
}

function New-NtnxClusterOverviewRow {
    param(
        [Parameter(Mandatory)]$Cluster,
        [hashtable]$AlertCounts = @{},
        [string]$PcName
    )

    $ft = Get-Prop $Cluster 'config.faultToleranceState' $null
    $faultTolerance = 'NA'
    $redundancyStatus = 'NA'
    if ($ft) {
        $v = Get-Prop $ft 'currentClusterFaultTolerance' (Get-Prop $ft 'currentMaxFaultTolerance' $null)
        if ($null -ne $v) { $faultTolerance = [string]$v }

        $rs = Get-Prop $ft 'redundancyStatus' $null
        if ($rs) {
            $pairs = foreach ($p in $rs.PSObject.Properties) {
                if ($p.Name -like '$*') { continue }
                if ($null -ne $p.Value) { '{0}={1}' -f $p.Name, $p.Value }
            }
            if ($pairs) { $redundancyStatus = ($pairs -join ', ') }
        }
    }

    $critical = 0; $warning = 0; $info = 0
    if ($AlertCounts.ContainsKey('CRITICAL')) { $critical = $AlertCounts['CRITICAL'] }
    if ($AlertCounts.ContainsKey('WARNING'))  { $warning  = $AlertCounts['WARNING'] }
    if ($AlertCounts.ContainsKey('INFO'))     { $info     = $AlertCounts['INFO'] }

    [ordered]@{
        'Cluster Name'         = Get-PropText $Cluster 'name'
        'Cluster UUID'         = Get-PropText $Cluster 'extId'
        'Cluster Function'     = Get-PropText $Cluster 'config.clusterFunction'
        'AOS Version'          = Get-PropText $Cluster 'config.buildInfo.version'
        'AOS Build'            = Get-PropText $Cluster 'config.buildInfo.fullVersion'
        'Hypervisor'           = Get-PropText $Cluster 'config.hypervisorTypes'
        'Number of Nodes'      = Get-Prop $Cluster 'nodes.numberOfNodes' 'NA'
        'Cluster Virtual IP'   = Get-NtnxIpValue (Get-Prop $Cluster 'network.externalAddress' $null)
        'Data Services IP'     = Get-NtnxIpValue (Get-Prop $Cluster 'network.externalDataServiceIp' $null)
        'Redundancy Factor'    = Get-Prop $Cluster 'config.redundancyFactor' 'NA'
        'Cluster FT (Current)' = $faultTolerance
        'Domain Awareness'     = Get-PropText $Cluster 'config.faultToleranceState.domainAwarenessLevel'
        'Redundancy Status'    = $redundancyStatus
        'Operation Mode'       = Get-PropText $Cluster 'config.operationMode'
        'LTS'                  = Get-PropBoolText $Cluster 'config.isLts'
        'Pulse'                = $(if (Get-Prop $Cluster 'config.pulseStatus.isEnabled' $false) { 'Enabled' } else { 'Disabled' })
        'Encryption (Transit)' = Get-PropText $Cluster 'config.encryptionInTransitStatus'
        'Remote Support'       = Get-PropBoolText $Cluster 'config.isRemoteSupportEnabled'
        'Timezone'             = Get-PropText $Cluster 'config.timezone'
        'VM Count'             = Get-Prop $Cluster 'vmCount' 'NA'
        'Inefficient VMs'      = Get-Prop $Cluster 'inefficientVmCount' 'NA'
        'Container Name'       = Get-PropText $Cluster 'containerName'
        'Critical Alerts'      = $critical
        'Warning Alerts'       = $warning
        'Info Alerts'          = $info
        'Total Open Alerts'    = $critical + $warning + $info
        'PC Name'              = $PcName
    }
}

function Get-NtnxClusterCategoryMap {
    <# Cluster-level categories flattened to a key -> value hashtable. #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster
    )

    $map = @{}
    $categories = Get-Prop $Cluster 'categories' $null
    if (-not $categories) { return $map }

    foreach ($entry in @($categories)) {
        $extId = if ($entry -is [string]) { $entry } else { Get-Prop $entry 'extId' $null }
        if (-not $extId) { continue }
        $formatted = Resolve-NtnxCategory -Session $Session -ExtId $extId
        if ($formatted -and $formatted.Contains(':')) {
            $parts = $formatted.Split(':', 2)
            $map[$parts[0]] = $parts[1]
        }
    }
    return $map
}

# --- Hosts ------------------------------------------------------------------------

function Get-NtnxHostRows {
    <#
    .SYNOPSIS
        Host rows plus the hostExtId -> hostName map used by the VM sheet.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster,
        [Parameter(Mandatory)][hashtable]$Window,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $hostMap = @{}

    $clusterExtId = Get-Prop $Cluster 'extId' $null
    $clusterName = Get-PropText $Cluster 'name'
    $listPath = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix "config/clusters/$clusterExtId/hosts"
    if (-not $listPath) { return @{ Rows = $rows; HostMap = $hostMap } }

    $hosts = @(Get-NtnxList -Session $Session -Path $listPath -OrderBy 'hostName')

    foreach ($h in $hosts) {
        try {
            $hostExtId = Get-Prop $h 'extId' $null
            $hostName = Get-PropText $h 'hostName'
            if ($hostExtId) { $hostMap[$hostExtId] = $hostName }

            # Per-host stats: clustermgmt exposes no list_host_stats, only this
            # per-host call. Host counts are small, so serial is fine.
            $stats = $null
            $statsPath = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix "stats/clusters/$clusterExtId/hosts/$hostExtId"
            if ($statsPath) {
                try {
                    $statsResp = Invoke-NtnxApi -Session $Session -Path $statsPath -Quiet -Query @{
                        '$startTime'        = $Window.Start
                        '$endTime'          = $Window.End
                        '$samplingInterval' = $script:SamplingInterval
                        '$statType'         = $script:StatType
                        '$select'           = '*'
                    }
                    $stats = Get-Prop $statsResp 'data' $null
                }
                catch { $stats = $null }
            }

            $memoryCapacityGib = ConvertTo-Gib (Get-Prop $h 'memorySizeBytes' 0)
            $cpuUsagePpm = Get-NtnxAverage (Get-Prop $stats 'hypervisorCpuUsagePpm' $null)
            $memoryUsedGib = ConvertTo-Gib (Get-NtnxAverage (Get-Prop $stats 'overallMemoryUsageBytes' $null))
            $storageCapacityGib = ConvertTo-Gib (Get-NtnxAverage (Get-Prop $stats 'storageCapacityBytes' $null)) -Floor
            $freeStorageGib = ConvertTo-Gib (Get-NtnxAverage (Get-Prop $stats 'freePhysicalStorageBytes' $null)) -Floor

            $sockets = Get-Prop $h 'numberOfCpuSockets' 0
            $cores = Get-Prop $h 'numberOfCpuCores' 0
            $coresPerSocket = 0
            if ($sockets) { $coresPerSocket = [math]::Round($cores / $sockets) }

            # NIC count is not always inlined on the host model.
            $nics = Get-Prop $h 'hostNics' $null
            $nicCount = @($nics).Count
            if (-not $nicCount) {
                $nicPath = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix "config/clusters/$clusterExtId/hosts/$hostExtId/host-nics"
                if ($nicPath) {
                    $nicCount = @(Get-NtnxList -Session $Session -Path $nicPath).Count
                }
            }

            $memoryUsagePercent = 0
            if ($memoryCapacityGib) { $memoryUsagePercent = [math]::Round($memoryUsedGib / $memoryCapacityGib * 100) }

            $rows.Add([ordered]@{
                'Host Name'              = $hostName
                'Cluster Name'           = $clusterName
                'Hypervisor Version'     = Get-PropText $h 'hypervisor.fullName'
                'Hypervisor IP'          = Get-NtnxIpValue (Get-Prop $h 'hypervisor.externalAddress' $null)
                'Host State'             = Get-PropText $h 'hypervisor.state'
                'CVM IP'                 = Get-NtnxIpValue (Get-Prop $h 'controllerVm.externalAddress' $null)
                'Model'                  = Get-PropText $h 'blockModel'
                'Host Serial No.'        = Get-PropText $h 'blockSerial'
                'CPU Model'              = Get-PropText $h 'cpuModel'
                "Physical CPU's"         = $sockets
                'Cores Per Socket'       = $coresPerSocket
                'Total CPU Cores'        = $cores
                'CPU Threads'            = Get-Prop $h 'numberOfCpuThreads' 0
                'CPU Usage %'            = [math]::Round($cpuUsagePpm / 10000, 2)
                'Total Memory (GiB)'     = [math]::Round($memoryCapacityGib)
                'Memory Usage %'         = $memoryUsagePercent
                'Storage Capacity (GiB)' = [math]::Round($storageCapacityGib)
                'Free Space (GiB)'       = [math]::Round($freeStorageGib)
                'NIC Count'              = $nicCount
                'Active VMs'             = Get-Prop $h 'hypervisor.numberOfVms' 0
                'PC Name'                = $PcName
            })
        }
        catch {
            Write-NtnxLog "Failed to process host $(Get-PropText $h 'hostName'): $($_.Exception.Message)" -Level WARN
        }
    }

    return @{ Rows = $rows; HostMap = $hostMap }
}

function Add-NtnxClusterCategoriesToHosts {
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $HostRows,
        [hashtable]$CategoryMap = @{}
    )

    foreach ($row in $HostRows) {
        $row['Environment'] = $CategoryMap[$script:ClusterEnvKey]
        $row['Location']    = $CategoryMap[$script:LocationKey]
        $row['Datacenter']  = $CategoryMap[$script:DatacenterKey]
        $row['Region']      = $CategoryMap[$script:RegionKey]
    }
}

# --- VMs --------------------------------------------------------------------------

function Get-NtnxVmStatsMap {
    <#
    .SYNOPSIS
        Fetch per-VM stats for powered-on VMs.
    .DESCRIPTION
        PowerShell 7 uses a throttled parallel pool; Windows PowerShell 5.1 has
        no equivalent and falls back to a serial loop. The throttle deliberately
        stays low (3) because Prism Central rate-limits these endpoints and more
        workers reduces throughput.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        $Vms,
        [Parameter(Mandatory)][hashtable]$Window,
        [string]$ClusterName
    )

    $map = @{}
    $targets = @($Vms | Where-Object {
        $ps = [string](Get-Prop $_ 'powerState' '')
        (Get-Prop $_ 'extId' $null) -and ($ps.ToUpper() -notlike '*OFF*')
    })

    $skipped = @($Vms).Count - $targets.Count
    if (-not $targets.Count) {
        if ($skipped) {
            Write-NtnxLog "  $ClusterName`: all $skipped VM(s) powered off - skipped stats fetch." -Level INFO
        }
        return $map
    }

    $statsBase = Get-NtnxPath -Session $Session -Namespace 'vmm' -Suffix 'ahv/stats/vms'
    if (-not $statsBase) { return $map }

    $workers = [math]::Min($script:VmStatsWorkers, $targets.Count)
    Write-NtnxLog "  $ClusterName`: fetching stats for $($targets.Count) powered-on VM(s) (concurrency=$workers, $skipped powered-off skipped) ..." -Level INFO

    $query = '?$startTime={0}&$endTime={1}&$samplingInterval={2}&$statType={3}&$select=*' -f `
        [uri]::EscapeDataString($Window.Start),
        [uri]::EscapeDataString($Window.End),
        $script:SamplingInterval,
        $script:StatType

    if ($Session.IsCore) {
        $baseUrl = $Session.BaseUrl
        $headers = $Session.Headers
        $timeout = $Session.TimeoutSec
        $extIds = $targets | ForEach-Object { Get-Prop $_ 'extId' $null }

        $results = $extIds | ForEach-Object -ThrottleLimit $workers -Parallel {
            $extId = $_
            $base = $using:baseUrl
            $statsPath = $using:statsBase
            $queryString = $using:query
            $requestHeaders = $using:headers
            $requestTimeout = $using:timeout

            try {
                $resp = Invoke-RestMethod -Method Get -Uri "$base/$statsPath/$extId$queryString" `
                    -Headers $requestHeaders -TimeoutSec $requestTimeout -SkipCertificateCheck -ErrorAction Stop
                [PSCustomObject]@{ ExtId = $extId; Stats = $resp.data.stats }
            }
            catch {
                [PSCustomObject]@{ ExtId = $extId; Stats = $null }
            }
        }

        foreach ($r in $results) { $map[$r.ExtId] = $r.Stats }
    }
    else {
        $done = 0
        foreach ($vm in $targets) {
            $extId = Get-Prop $vm 'extId' $null
            try {
                $resp = Invoke-NtnxApi -Session $Session -Path "$statsBase/$extId" -Quiet -Query @{
                    '$startTime'        = $Window.Start
                    '$endTime'          = $Window.End
                    '$samplingInterval' = $script:SamplingInterval
                    '$statType'         = $script:StatType
                    '$select'           = '*'
                }
                $map[$extId] = Get-Prop $resp 'data.stats' $null
            }
            catch {
                $map[$extId] = $null
            }
            $done++
            if ($done % 25 -eq 0) {
                Write-NtnxLog "    stats $done/$($targets.Count) ..." -Level DEBUG
            }
        }
    }

    return $map
}

function Get-NtnxVmRows {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster,
        [Parameter(Mandatory)][hashtable]$Window,
        [hashtable]$HostMap = @{},
        [double]$UnderProvisionPercent = 75.0,
        [double]$OverProvisionPercent = 30.0,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $clusterExtId = Get-Prop $Cluster 'extId' $null
    $clusterName = Get-PropText $Cluster 'name'

    $vmPath = Get-NtnxPath -Session $Session -Namespace 'vmm' -Suffix 'ahv/config/vms'
    if (-not $vmPath) { return $rows }

    $vms = @(Get-NtnxList -Session $Session -Path $vmPath -OrderBy 'name' -Filter "cluster/extId eq '$clusterExtId'")
    if (-not $vms.Count) { return $rows }

    $statsMap = Get-NtnxVmStatsMap -Session $Session -Vms $vms -Window $Window -ClusterName $clusterName

    foreach ($vm in $vms) {
        try {
            $extId = Get-Prop $vm 'extId' $null

            $vcpu = (Get-Prop $vm 'numSockets' 0) * (Get-Prop $vm 'numCoresPerSocket' 0) * (Get-Prop $vm 'numThreadsPerCore' 0)
            $memBytes = Get-Prop $vm 'memorySizeBytes' 0
            $memGib = [math]::Round((ConvertTo-Gib $memBytes))

            # IPs can appear as a static config or as learned addresses.
            $ips = [System.Collections.Generic.HashSet[string]]::new()
            $nics = @(Get-Prop $vm 'nics' @())
            foreach ($nic in $nics) {
                $configured = Get-Prop $nic 'networkInfo.ipv4Config.ipAddress.value' $null
                if ($configured) { [void]$ips.Add([string]$configured) }

                foreach ($learned in @(Get-Prop $nic 'networkInfo.ipv4Info.learnedIpAddresses' @())) {
                    $v = Get-NtnxIpValue $learned
                    if ($v -and $v -ne 'NA') { [void]$ips.Add($v) }
                }
                foreach ($addr in @(Get-Prop $nic 'ipv4Addresses' @())) {
                    $v = Get-NtnxIpValue $addr
                    if ($v -and $v -ne 'NA') { [void]$ips.Add($v) }
                }
            }

            $diskBytes = 0
            foreach ($disk in @(Get-Prop $vm 'disks' @())) {
                $diskBytes += [double](Get-Prop $disk 'backingInfo.diskSizeBytes' 0)
            }

            $categoryList = foreach ($category in @(Get-Prop $vm 'categories' @())) {
                $catExtId = if ($category -is [string]) { $category } else { Get-Prop $category 'extId' $null }
                if (-not $catExtId) { continue }
                $formatted = Resolve-NtnxCategory -Session $Session -ExtId $catExtId
                if ($formatted -and $formatted -notlike '*CalmProject*') { $formatted }
            }

            $stats = @()
            if ($extId -and $statsMap.ContainsKey($extId) -and $statsMap[$extId]) { $stats = @($statsMap[$extId]) }

            $cpuPpm = Get-NtnxAverage ($stats | ForEach-Object { Get-Prop $_ 'hypervisorCpuUsagePpm' $null } | Where-Object { $null -ne $_ })
            $memPpm = Get-NtnxAverage ($stats | ForEach-Object { Get-Prop $_ 'memoryUsagePpm' $null } | Where-Object { $null -ne $_ })
            $diskPpm = Get-NtnxAverage ($stats | ForEach-Object { Get-Prop $_ 'diskUsagePpm' $null } | Where-Object { $null -ne $_ })

            $cpuPercent = if ($cpuPpm) { [math]::Round($cpuPpm / 10000) } else { 0 }
            $memPercent = if ($memPpm) { [math]::Round($memPpm / 10000) } else { 0 }
            $diskPercent = if ($diskPpm) { [math]::Round($diskPpm / 10000) } else { 0 }

            $cpuColor = $script:ColorAdequate
            if ($cpuPercent -ne 0) {
                if ($cpuPercent -lt $OverProvisionPercent) { $cpuColor = $script:ColorOverProvisioned }
                elseif ($cpuPercent -gt $UnderProvisionPercent) { $cpuColor = $script:ColorUnderProvisioned }
            }
            $memColor = $script:ColorAdequate
            if ($memPercent -ne 0) {
                if ($memPercent -lt $OverProvisionPercent) { $memColor = $script:ColorOverProvisioned }
                elseif ($memPercent -gt $UnderProvisionPercent) { $memColor = $script:ColorUnderProvisioned }
            }

            $consumedMemGib = 0
            if ($memPpm) { $consumedMemGib = [math]::Round(($memPpm * $memBytes) / 1000000 / [math]::Pow(1024, 3), 2) }

            $hostExtId = Get-Prop $vm 'host.extId' $null
            $hostName = 'NA'
            if ($hostExtId -and $HostMap.ContainsKey($hostExtId)) { $hostName = $HostMap[$hostExtId] }

            $rows.Add([ordered]@{
                'VM Name'               = Get-PropText $vm 'name'
                'Parent Cluster'        = $clusterName
                'On Host'               = $hostName
                'Creation Time'         = Get-Prop $vm 'createTime' $null
                'Last Update Time'      = Get-Prop $vm 'updateTime' $null
                'UUID'                  = $extId
                'Power State'           = Get-PropText $vm 'powerState'
                'vCPU'                  = $vcpu
                'vCPU usage %'          = $(if ($cpuPercent -ne 0) { $cpuPercent } else { 'NA' })
                'CPU Usage Color'       = $cpuColor
                'Memory (GiB)'          = $memGib
                'Memory usage %'        = $(if ($memPercent -ne 0) { $memPercent } else { 'NA' })
                'Memory Usage Color'    = $memColor
                'Consumed Memory (GiB)' = $consumedMemGib
                'Disk Space (GiB)'      = [math]::Round((ConvertTo-Gib $diskBytes))
                'Disk Usage %'          = $(if ($diskPercent) { $diskPercent } else { 'NA' })
                'vNIC'                  = $nics.Count
                'IP Address'            = (($ips | Sort-Object) -join ',')
                'OS'                    = Get-PropText $vm 'guestTools.guestOsVersion'
                'NGT Installed'         = Get-PropBoolText $vm 'guestTools.isInstalled'
                'NGT Version'           = Get-PropText $vm 'guestTools.version'
                'Memory Overcommit'     = $(if (Get-Prop $vm 'isMemoryOvercommitEnabled' $false) { 'Enabled' } else { 'Disabled' })
                'Protection Status'     = Get-PropText $vm 'protectionType'
                'Categories'            = ($categoryList -join ',')
                'PC Name'               = $PcName
            })
        }
        catch {
            Write-NtnxLog "Error processing VM $(Get-PropText $vm 'name'): $($_.Exception.Message)" -Level WARN
        }
    }

    return $rows
}

# --- Storage containers -----------------------------------------------------------

function Get-NtnxStorageRows {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $clusterExtId = Get-Prop $Cluster 'extId' $null
    $clusterName = Get-PropText $Cluster 'name'

    $path = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix 'config/storage-containers'
    if (-not $path) { return $rows }

    $containers = Get-NtnxList -Session $Session -Path $path -OrderBy 'name asc' -Filter "clusterExtId eq '$clusterExtId'"

    foreach ($c in $containers) {
        $rows.Add([ordered]@{
            'Container Name'            = Get-PropText $c 'name'
            'Cluster'                   = $clusterName
            'PC Name'                   = $PcName
            'Replication Factor'        = Get-Prop $c 'replicationFactor' 'NA'
            'Compression'               = Get-PropBoolText $c 'isCompressionEnabled'
            'Erasure Coding'            = Get-PropText $c 'erasureCode'
            'Encrypted'                 = Get-PropBoolText $c 'isEncrypted'
            'Internal'                  = Get-PropBoolText $c 'isInternal'
            'Max Capacity (GiB)'        = ConvertTo-Gib (Get-Prop $c 'maxCapacityBytes' 0)
            'Advertised Capacity (GiB)' = ConvertTo-Gib (Get-Prop $c 'logicalAdvertisedCapacityBytes' 0)
            'Container UUID'            = Get-PropText $c 'extId' (Get-PropText $c 'containerExtId')
        })
    }
    return $rows
}

# --- Host disks -------------------------------------------------------------------

function Get-NtnxDiskRows {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $clusterExtId = Get-Prop $Cluster 'extId' $null
    $clusterName = Get-PropText $Cluster 'name'

    $path = Get-NtnxPath -Session $Session -Namespace 'clustermgmt' -Suffix 'config/disks'
    if (-not $path) { return $rows }

    $disks = Get-NtnxList -Session $Session -Path $path -Filter "clusterExtId eq '$clusterExtId'"

    foreach ($d in $disks) {
        $rows.Add([ordered]@{
            'Cluster'             = $clusterName
            'Host Name'           = Get-PropText $d 'hostName'
            'Host IP'             = Get-NtnxIpValue (Get-Prop $d 'nodeIpAddress' $null)
            'CVM IP'              = Get-NtnxIpValue (Get-Prop $d 'cvmIpAddress' $null)
            'Tier'                = Get-PropText $d 'storageTier'
            'Status'              = Get-PropText $d 'status'
            'Model'               = Get-PropText $d 'model'
            'Vendor'              = Get-PropText $d 'vendor'
            'Serial Number'       = Get-PropText $d 'serialNumber'
            'Mount Path'          = Get-PropText $d 'mountPath'
            'Logical Size (GiB)'  = ConvertTo-Gib (Get-Prop $d 'diskSizeBytes' 0)
            'Physical Size (GiB)' = ConvertTo-Gib (Get-Prop $d 'physicalCapacityBytes' 0)
            'Firmware'            = Get-PropText $d 'firmwareVersion'
            'Disk UUID'           = Get-PropText $d 'extId'
            'PC Name'             = $PcName
        })
    }
    return $rows
}

# --- Subnets / virtual switches ---------------------------------------------------

function Get-NtnxSubnetRows {
    <#
        Listed PC-wide rather than per-cluster: some PCs return subnets without a
        populated clusterReference, so a filtered query yields nothing even when
        subnets exist. Each row is mapped back to its cluster afterwards.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [hashtable]$UuidToName = @{},
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $path = Get-NtnxPath -Session $Session -Namespace 'networking' -Suffix 'config/subnets'
    if (-not $path) { return $rows }

    foreach ($s in Get-NtnxList -Session $Session -Path $path -OrderBy 'name') {
        $clusterRef = Get-Prop $s 'clusterReference' $null
        $clusterExtId = if ($clusterRef -is [string]) { $clusterRef } else { Get-Prop $clusterRef 'extId' $null }
        $clusterName = 'NA'
        if ($clusterExtId -and $UuidToName.ContainsKey($clusterExtId)) { $clusterName = $UuidToName[$clusterExtId] }
        elseif (Get-Prop $s 'clusterName' $null) { $clusterName = Get-PropText $s 'clusterName' }

        $ipSummary = foreach ($ic in @(Get-Prop $s 'ipConfig' @())) {
            $v4 = Get-Prop $ic 'ipv4' $null
            if (-not $v4) { continue }
            $ip = Get-Prop $v4 'ipSubnet' (Get-Prop $v4 'defaultGatewayIp' $null)
            $value = Get-NtnxIpValue $ip
            if ($value -and $value -ne 'NA') { $value }
        }

        $vsRef = Get-Prop $s 'virtualSwitchReference' (Get-Prop $s 'virtualSwitch' $null)
        $vsText = if ($vsRef -is [string]) { $vsRef } elseif ($vsRef) { Get-PropText $vsRef 'extId' } else { 'NA' }

        $rows.Add([ordered]@{
            'Cluster'              = $clusterName
            'Subnet Name'          = Get-PropText $s 'name'
            'Type'                 = Get-PropText $s 'subnetType'
            'VLAN ID'              = Get-Prop $s 'networkId' 'NA'
            'Virtual Switch'       = $vsText
            'Bridge'               = Get-PropText $s 'bridgeName'
            'Hypervisor Type'      = Get-PropText $s 'hypervisorType'
            'External'             = Get-PropBoolText $s 'isExternal'
            'NAT'                  = Get-PropBoolText $s 'isNatEnabled'
            'Advanced Networking'  = Get-PropBoolText $s 'isAdvancedNetworking'
            'IP Config'            = $(if ($ipSummary) { ($ipSummary -join ', ') } else { 'NA' })
            'Free IPs (Pool)'      = Get-Prop $s 'ipUsage.numFreeIpsInPool' 'NA'
            'Assigned IPs'         = Get-Prop $s 'ipUsage.numAssignedIps' 'NA'
            'Subnet UUID'          = Get-PropText $s 'extId'
            'PC Name'              = $PcName
        })
    }
    return $rows
}

function Get-NtnxVirtualSwitchRows {
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)]$Cluster,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $path = Get-NtnxPath -Session $Session -Namespace 'networking' -Suffix 'config/virtual-switches'
    if (-not $path) { return $rows }

    $clusterExtId = Get-Prop $Cluster 'extId' $null
    $clusterName = Get-PropText $Cluster 'name'

    $switches = Get-NtnxList -Session $Session -Path $path -OrderBy 'name' -ExtraHeaders @{ 'X-Cluster-Id' = $clusterExtId }

    foreach ($vs in $switches) {
        $rows.Add([ordered]@{
            'Cluster'            = $clusterName
            'Switch Name'        = Get-PropText $vs 'name'
            'Description'        = Get-PropText $vs 'description'
            'Bond Mode'          = Get-PropText $vs 'bondMode'
            'MTU'                = Get-Prop $vs 'mtu' 'NA'
            'Default'            = Get-PropBoolText $vs 'isDefault'
            'Quick Mode'         = Get-PropBoolText $vs 'isQuickMode'
            'Deployment Error'   = Get-PropBoolText $vs 'hasDeploymentError'
            'Update In Progress' = Get-PropBoolText $vs 'hasUpdateInProgress'
            'Owner Type'         = Get-PropText $vs 'ownerType'
            'Linked Clusters'    = @(Get-Prop $vs 'clusters' @()).Count
            'Switch UUID'        = Get-PropText $vs 'extId'
            'PC Name'            = $PcName
        })
    }
    return $rows
}

# --- Recovery points --------------------------------------------------------------

function Get-NtnxRecoveryPointRows {
    <#
        The data-protection API expects X-Cluster-Id to be a cluster *type* enum
        (AOS/MST), not a cluster UUID, so everything is listed once and mapped
        back to a cluster afterwards.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [hashtable]$UuidToName = @{},
        [string]$PcName,
        [int]$MaxRows = 1000
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $path = Get-NtnxPath -Session $Session -Namespace 'dataprotection' -Suffix 'config/recovery-points'
    if (-not $path) { return $rows }

    $points = Get-NtnxList -Session $Session -Path $path -OrderBy 'creationTime desc' `
        -ExtraHeaders @{ 'X-Cluster-Id' = 'AOS' } -MaxRows $MaxRows

    foreach ($rp in $points) {
        $clusterName = 'NA'
        foreach ($loc in @(Get-Prop $rp 'locationReferences' @())) {
            $ext = Get-Prop $loc 'locationExtId' $null
            if ($ext -and $UuidToName.ContainsKey($ext)) { $clusterName = $UuidToName[$ext]; break }
        }
        if ($clusterName -eq 'NA') {
            foreach ($vmRp in @(Get-Prop $rp 'vmRecoveryPoints' @())) {
                $ext = Get-Prop $vmRp 'clusterExtId' (Get-Prop $vmRp 'sourceClusterExtId' $null)
                if (-not $ext) { $ext = Get-Prop $vmRp 'locationAgnosticId' $null }
                if ($ext -and $UuidToName.ContainsKey($ext)) { $clusterName = $UuidToName[$ext]; break }
            }
        }

        $rows.Add([ordered]@{
            'Cluster'               = $clusterName
            'Name'                  = Get-PropText $rp 'name'
            'Type'                  = Get-PropText $rp 'recoveryPointType'
            'Status'                = Get-PropText $rp 'status'
            'Created'               = Get-Prop $rp 'creationTime' $null
            'Expires'               = Get-Prop $rp 'expirationTime' $null
            'Exclusive Usage (GiB)' = ConvertTo-Gib (Get-Prop $rp 'totalExclusiveUsageBytes' 0) -Digits 3
            'Owner ExtId'           = Get-PropText $rp 'ownerExtId'
            'Recovery Point UUID'   = Get-PropText $rp 'extId'
            'PC Name'               = $PcName
        })
    }
    return $rows
}

# --- Alerts -----------------------------------------------------------------------

function Get-NtnxAlertData {
    <#
        Returns open alerts (critical/warning rows for the workbook) plus a
        per-cluster severity tally used by the Cluster Overview sheet.

        clusterUuid is not a filterable field on this API, so we filter only on
        isResolved and resolve the owning cluster afterwards - this also picks up
        PC-originated alerts.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [hashtable]$UuidToName = @{},
        [string]$PcName,
        [int]$MaxRows = 1500
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $countsByCluster = @{}

    $path = Get-NtnxPath -Session $Session -Namespace 'monitoring' -Suffix 'serviceability/alerts'
    if (-not $path) { return @{ Rows = $rows; Counts = $countsByCluster } }

    $alerts = Get-NtnxList -Session $Session -Path $path -Filter 'isResolved eq false' `
        -OrderBy 'creationTime desc' -MaxRows $MaxRows

    foreach ($a in $alerts) {
        $severity = (Get-PropText $a 'severity').ToUpper()

        $clusterName = 'Prism Central'
        foreach ($attr in @('originatingClusterUuid', 'clusterUuid', 'originatingClusterExtId')) {
            $ext = Get-Prop $a $attr $null
            if ($ext -and $UuidToName.ContainsKey($ext)) { $clusterName = $UuidToName[$ext]; break }
        }

        if (-not $countsByCluster.ContainsKey($clusterName)) { $countsByCluster[$clusterName] = @{} }
        $bucket = $countsByCluster[$clusterName]
        if (-not $bucket.ContainsKey($severity)) { $bucket[$severity] = 0 }
        $bucket[$severity]++

        if ($severity -in @('CRITICAL', 'WARNING')) {
            $rows.Add([ordered]@{
                'Cluster'      = $clusterName
                'Severity'     = $severity
                'Title'        = Get-PropText $a 'title'
                'Message'      = Get-PropText $a 'message'
                'Service'      = Get-PropText $a 'serviceName'
                'Alert Type'   = Get-PropText $a 'alertType'
                'Created'      = Get-Prop $a 'creationTime' $null
                'Last Updated' = Get-Prop $a 'lastUpdatedTime' $null
                'Acknowledged' = Get-PropBoolText $a 'isAcknowledged'
                'Auto Resolved' = Get-PropBoolText $a 'isAutoResolved'
                'Alert UUID'   = Get-PropText $a 'extId'
                'PC Name'      = $PcName
            })
        }
    }

    return @{ Rows = $rows; Counts = $countsByCluster }
}

# --- Licenses ---------------------------------------------------------------------

function Get-NtnxLicenseRows {
    param(
        [Parameter(Mandatory)]$Session,
        [string]$PcName
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $path = Get-NtnxPath -Session $Session -Namespace 'licensing' -Suffix 'config/licenses'
    if (-not $path) { return $rows }

    foreach ($lic in Get-NtnxList -Session $Session -Path $path) {
        $rows.Add([ordered]@{
            'Name'          = Get-PropText $lic 'name'
            'Category'      = Get-PropText $lic 'category'
            'Sub-category'  = Get-PropText $lic 'subCategory'
            'Type'          = Get-PropText $lic 'type'
            'Scope'         = Get-PropText $lic 'scope'
            'Quantity'      = Get-Prop $lic 'quantity' 'NA'
            'Meter'         = Get-PropText $lic 'meter'
            'Expiry Date'   = Get-Prop $lic 'expiryDate' $null
            'Salesforce ID' = Get-PropText $lic 'salesforceLicenseId'
            'License UUID'  = Get-PropText $lic 'extId'
            'PC Name'       = $PcName
        })
    }
    return $rows
}

# --- Protection policies ----------------------------------------------------------

function ConvertTo-HumanSeconds {
    param($Seconds)

    $value = 0
    if (-not [int]::TryParse([string]$Seconds, [ref]$value)) { return 'NA' }
    if ($value -le 0) { return 'NA' }

    $parts = @()
    foreach ($unit in @(@(86400, 'd'), @(3600, 'h'), @(60, 'm'), @(1, 's'))) {
        $size = $unit[0]
        if ($value -ge $size) {
            $count = [math]::Floor($value / $size)
            $parts += "$count$($unit[1])"
            $value = $value % $size
        }
    }
    if (-not $parts) { return 'NA' }
    return ($parts -join ' ')
}

function ConvertTo-HumanRetention {
    param($Retention)

    if (-not $Retention) { return 'NA' }

    foreach ($path in @('localSnapshotRetentionCount', 'snapshotRetentionCount', 'count', 'quantity')) {
        $v = Get-Prop $Retention $path $null
        if ($v) { return "$v snapshot(s)" }
    }
    foreach ($path in @('retentionTimeSeconds', 'localRetentionTimeSeconds', 'timeSeconds')) {
        $v = Get-Prop $Retention $path $null
        if ($v) { return ConvertTo-HumanSeconds $v }
    }
    return 'NA'
}

function Get-NtnxProtectionPolicyRows {
    param(
        [Parameter(Mandatory)]$Session,
        [hashtable]$UuidToName = @{},
        [string]$PcName,
        [int]$MaxRows = 500
    )

    $rows = [System.Collections.Generic.List[object]]::new()
    $path = Get-NtnxPath -Session $Session -Namespace 'datapolicies' -Suffix 'config/protection-policies'
    if (-not $path) { return $rows }

    foreach ($p in Get-NtnxList -Session $Session -Path $path -OrderBy 'name' -MaxRows $MaxRows) {
        $locations = @(Get-Prop $p 'replicationLocations' @())
        $primaryLabel = 'NA'
        $primaryCluster = 'NA'
        $labels = @()

        foreach ($loc in $locations) {
            $label = Get-PropText $loc 'label' 'site'
            $labels += $label
            if (Get-Prop $loc 'isPrimary' $false) {
                $primaryLabel = $label
                $sub = Get-Prop $loc 'replicationSubLocation' $null
                if ($sub) {
                    $refs = Get-Prop $sub 'clusterExtIds' (Get-Prop $sub 'nutanixCluster.clusterExtIds' $null)
                    $first = @($refs) | Select-Object -First 1
                    if ($first) {
                        $primaryCluster = if ($UuidToName.ContainsKey($first)) { $UuidToName[$first] } else { [string]$first }
                    }
                }
            }
        }

        $rpoText = 'NA'; $retentionText = 'NA'; $rpType = 'NA'
        $scheduleCount = 0; $syncCount = 0; $asyncCount = 0; $crossAz = $false

        foreach ($cfg in @(Get-Prop $p 'replicationConfigurations' @())) {
            $scheduleCount++
            $sched = Get-Prop $cfg 'schedule' $null
            if (-not $sched) { continue }

            $rpoSecs = Get-Prop $sched 'recoveryPointObjectiveTimeSeconds' $null
            if ($null -ne $rpoSecs) {
                if ($rpoSecs -eq 0) { $syncCount++ } else { $asyncCount++ }
                if ($rpoText -eq 'NA' -and $rpoSecs) { $rpoText = ConvertTo-HumanSeconds $rpoSecs }
            }

            if ($rpType -eq 'NA') { $rpType = Get-PropText $sched 'recoveryPointType' }
            if ($retentionText -eq 'NA') { $retentionText = ConvertTo-HumanRetention (Get-Prop $sched 'retention' $null) }

            $src = Get-Prop $cfg 'sourceLocationLabel' $null
            $remote = Get-Prop $cfg 'remoteLocationLabel' $null
            if ($src -and $remote -and $src -ne $remote) { $crossAz = $true }
        }

        $rows.Add([ordered]@{
            'Policy Name'           = Get-PropText $p 'name'
            'Description'           = Get-PropText $p 'description'
            'Primary Site'          = $primaryLabel
            'Primary Cluster'       = $primaryCluster
            'Replication Locations' = $(if ($labels) { (($labels | Sort-Object -Unique) -join ', ') } else { 'NA' })
            'Locations'             = $locations.Count
            'Schedules'             = $scheduleCount
            'RPO'                   = $rpoText
            'Sync Schedules'        = $syncCount
            'Async Schedules'       = $asyncCount
            'Cross-AZ'              = $(if ($crossAz) { 'Yes' } else { 'No' })
            'Recovery Point Type'   = $rpType
            'Retention'             = $retentionText
            'Categories'            = Get-PropText $p 'categoryIds'
            'Approval Required'     = Get-PropBoolText $p 'isApprovalPolicyNeeded'
            'Owner ExtId'           = Get-PropText $p 'ownerExtId'
            'Policy UUID'           = Get-PropText $p 'extId'
            'PC Name'               = $PcName
        })
    }
    return $rows
}

# --- Orchestrator -----------------------------------------------------------------

function Get-NtnxPcInventory {
    <#
    .SYNOPSIS
        Collect the full inventory for one Prism Central.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PcIp,
        [Parameter(Mandatory)][string]$PcUser,
        [string]$PcSecret,
        [ValidateSet('basic', 'api_key')][string]$AuthMode = 'basic',
        [string[]]$ClusterNames = @('ALL'),
        [double]$UnderProvisionPercent = 75.0,
        [double]$OverProvisionPercent = 30.0,
        [string]$ApiPathVersion = 'auto'
    )

    $window = Get-NtnxUtcWindow
    $generatedAt = [DateTime]::UtcNow
    $timestampSlug = $generatedAt.ToString('yyyy-MM-dd_HH-mm-ss')

    $session = New-NtnxSession -PcIp $PcIp -User $PcUser -Secret $PcSecret -AuthMode $AuthMode -ApiPathVersion $ApiPathVersion

    Write-NtnxLog "Listing clusters on Prism Central $PcIp ..."
    $clusters = @(Get-NtnxClusters -Session $session)
    if (-not $clusters.Count) { throw "No clusters returned for PC $PcIp" }

    # The PC itself appears in the cluster list; identify it and exclude it.
    $pcName = ''
    $pcVersion = 'NA'
    $selected = [System.Collections.Generic.List[object]]::new()
    $uuidToName = @{}

    foreach ($c in $clusters) {
        $name = Get-PropText $c 'name'
        $extId = Get-Prop $c 'extId' $null
        if ($extId) { $uuidToName[$extId] = $name }

        $functions = @(Get-Prop $c 'config.clusterFunction' @()) | ForEach-Object { [string]$_ }
        if ($functions -contains 'PRISM_CENTRAL') {
            $pcName = $name
            $pcVersion = Get-PropText $c 'config.buildInfo.version'
            continue
        }
        if ($ClusterNames -and ($ClusterNames -notcontains 'ALL') -and ($ClusterNames -notcontains $name)) { continue }
        $selected.Add($c)
    }

    if (-not $pcName) { $pcName = 'PC_' + $PcIp.Replace('.', '_') }
    if (-not $selected.Count) { throw "No clusters selected to process for $PcIp" }

    Write-NtnxLog "Discovered $($clusters.Count) cluster(s); processing $($selected.Count) ($(($selected | ForEach-Object { Get-PropText $_ 'name' }) -join ', '))."

    Write-NtnxLog "Fetching PC-wide open alerts ..."
    $alertData = Get-NtnxAlertData -Session $session -UuidToName $uuidToName -PcName $pcName
    $alertRows = $alertData.Rows
    Write-NtnxLog "Open alerts: $($alertRows.Count) (critical+warning shown in workbook)."

    $allVms = [System.Collections.Generic.List[object]]::new()
    $allHosts = [System.Collections.Generic.List[object]]::new()
    $allStorage = [System.Collections.Generic.List[object]]::new()
    $allDisks = [System.Collections.Generic.List[object]]::new()
    $allSwitches = [System.Collections.Generic.List[object]]::new()
    $overviewRows = [System.Collections.Generic.List[object]]::new()
    $processed = @()

    $index = 0
    foreach ($cluster in $selected) {
        $index++
        $clusterName = Get-PropText $cluster 'name'
        Write-NtnxLog "[Cluster $index/$($selected.Count)] $clusterName - gathering inventory ..."
        $processed += $clusterName

        $hostResult = Get-NtnxHostRows -Session $session -Cluster $cluster -Window $window -PcName $pcName
        $categoryMap = Get-NtnxClusterCategoryMap -Session $session -Cluster $cluster
        Add-NtnxClusterCategoriesToHosts -HostRows $hostResult.Rows -CategoryMap $categoryMap
        foreach ($r in $hostResult.Rows) { $allHosts.Add($r) }

        $vmRows = @(Get-NtnxVmRows -Session $session -Cluster $cluster -Window $window -HostMap $hostResult.HostMap `
            -UnderProvisionPercent $UnderProvisionPercent -OverProvisionPercent $OverProvisionPercent -PcName $pcName)
        foreach ($r in $vmRows) { $allVms.Add($r) }

        $storageRows = @(Get-NtnxStorageRows -Session $session -Cluster $cluster -PcName $pcName)
        foreach ($r in $storageRows) { $allStorage.Add($r) }

        $diskRows = @(Get-NtnxDiskRows -Session $session -Cluster $cluster -PcName $pcName)
        foreach ($r in $diskRows) { $allDisks.Add($r) }

        $switchRows = @(Get-NtnxVirtualSwitchRows -Session $session -Cluster $cluster -PcName $pcName)
        foreach ($r in $switchRows) { $allSwitches.Add($r) }

        $counts = @{}
        if ($alertData.Counts.ContainsKey($clusterName)) { $counts = $alertData.Counts[$clusterName] }
        $overviewRows.Add((New-NtnxClusterOverviewRow -Cluster $cluster -AlertCounts $counts -PcName $pcName))

        Write-NtnxLog ("  {0}: {1} hosts, {2} VMs, {3} disks, {4} storage containers, {5} vSwitches" -f `
            $clusterName, $hostResult.Rows.Count, $vmRows.Count, $diskRows.Count, $storageRows.Count, $switchRows.Count)
    }

    Write-NtnxLog "Fetching PC-wide subnets, recovery points, protection policies, licenses ..."
    $subnetRows = @(Get-NtnxSubnetRows -Session $session -UuidToName $uuidToName -PcName $pcName)
    $rpRows = @(Get-NtnxRecoveryPointRows -Session $session -UuidToName $uuidToName -PcName $pcName)
    $policyRows = @(Get-NtnxProtectionPolicyRows -Session $session -UuidToName $uuidToName -PcName $pcName)
    $licenseRows = @(Get-NtnxLicenseRows -Session $session -PcName $pcName)
    Write-NtnxLog ("PC-wide totals -> subnets={0}, recovery points={1}, protection policies={2}, licenses={3}" -f `
        $subnetRows.Count, $rpRows.Count, $policyRows.Count, $licenseRows.Count)

    $severityOrder = @{ 'CRITICAL' = 0; 'WARNING' = 1; 'INFO' = 2 }

    $summary = [ordered]@{
        pc_name                 = $pcName
        pc_ip                   = $PcIp
        pc_version              = $pcVersion
        cluster_count           = $processed.Count
        vm_count                = $allVms.Count
        host_count              = $allHosts.Count
        disk_count              = $allDisks.Count
        storage_container_count = $allStorage.Count
        subnet_count            = $subnetRows.Count
        virtual_switch_count    = $allSwitches.Count
        recovery_point_count    = $rpRows.Count
        protection_policy_count = $policyRows.Count
        license_count           = $licenseRows.Count
        critical_alerts         = @($alertRows | Where-Object { $_['Severity'] -eq 'CRITICAL' }).Count
        warning_alerts          = @($alertRows | Where-Object { $_['Severity'] -eq 'WARNING' }).Count
    }

    return [PSCustomObject]@{
        PcName                = $pcName
        PcIp                  = $PcIp
        GeneratedAtUtc        = $generatedAt
        TimestampSlug         = $timestampSlug
        ClustersProcessed     = $processed
        ClusterOverviewRows   = @($overviewRows)
        VmRows                = @($allVms | Sort-Object { $_['Parent Cluster'] }, { $_['VM Name'] })
        HostRows              = @($allHosts | Sort-Object { $_['Cluster Name'] }, { $_['Host Name'] })
        HostDiskRows          = @($allDisks | Sort-Object { $_['Cluster'] }, { $_['Host Name'] }, { $_['Mount Path'] })
        StorageRows           = @($allStorage | Sort-Object { $_['Cluster'] }, { $_['Container Name'] })
        SubnetRows            = @($subnetRows | Sort-Object { $_['Cluster'] }, { $_['Subnet Name'] })
        VSwitchRows           = @($allSwitches | Sort-Object { $_['Cluster'] }, { $_['Switch Name'] })
        RecoveryPointRows     = @($rpRows | Sort-Object { $_['Created'] } -Descending)
        ProtectionPolicyRows  = @($policyRows | Sort-Object { $_['Policy Name'] })
        AlertRows             = @($alertRows | Sort-Object { $severityOrder[[string]$_['Severity']] }, { $_['Cluster'] })
        LicenseRows           = @($licenseRows)
        PcSummary             = $summary
    }
}
