<#
.SYNOPSIS
    Nutanix Prism Central inventory - produces one timestamped Excel workbook per PC.

.DESCRIPTION
    Collects clusters, hosts, host disks, VMs (with 24h CPU/memory/disk usage),
    storage containers, subnets, virtual switches, recovery points, protection
    policies, open alerts and licences from the Prism Central v4 REST API, and
    writes a formatted multi-sheet workbook.

    Three ways to run:
      1. Interactive          - .\Invoke-NtnxInventory.ps1
      2. Single / multiple PC - .\Invoke-NtnxInventory.ps1 -PcIp 10.0.0.1 -PcUser admin -PcSecret '...'
      3. Credentials file     - .\Invoke-NtnxInventory.ps1 -CredentialsFile .\pcs.csv

    Credentials may also come from the environment (PRISM_CENTRAL_VIP,
    PRISM_CENTRAL_USERNAME, PRISM_CENTRAL_PASSWORD) so the script can run
    unattended with secrets injected at execution time.

.PARAMETER PcIp
    One or more Prism Central IPs/FQDNs. Accepts 'host' or 'host:port'.

.PARAMETER AuthMode
    'basic' (default) uses PcUser/PcSecret. 'api_key' sends PcUser as the key.

.PARAMETER Clusters
    Cluster names to include. Defaults to ALL.

.PARAMETER CredentialsFile
    CSV or XLSX with columns pc_ip, pc_user and optionally pc_secret, auth_mode,
    clusters, api_path_version.

.PARAMETER ApiPathVersion
    Pin the v4 minor version (e.g. v4.0) instead of probing each namespace.

.EXAMPLE
    .\Invoke-NtnxInventory.ps1 -PcIp 10.38.99.39 -PcUser admin -PcSecret 'secret'

.EXAMPLE
    .\Invoke-NtnxInventory.ps1 -CredentialsFile .\pcs.csv -OutputPath D:\Reports
#>

[CmdletBinding(DefaultParameterSetName = 'Direct')]
param(
    [Parameter(ParameterSetName = 'Direct')]
    [string[]]$PcIp,

    [Parameter(ParameterSetName = 'Direct')]
    [string]$PcUser,

    [Parameter(ParameterSetName = 'Direct')]
    [string]$PcSecret,

    [Parameter(ParameterSetName = 'Direct')]
    [ValidateSet('basic', 'api_key')]
    [string]$AuthMode = 'basic',

    [Parameter(ParameterSetName = 'File', Mandatory)]
    [string]$CredentialsFile,

    [string[]]$Clusters = @('ALL'),

    [string]$OutputPath = (Join-Path $PSScriptRoot 'output'),

    [double]$UnderProvisionPercent = 75.0,

    [double]$OverProvisionPercent = 30.0,

    [string]$ApiPathVersion = 'auto',

    [switch]$Interactive
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# Three-argument Join-Path is PowerShell 6+, so nest the calls for 5.1.
$moduleRoot = Join-Path $PSScriptRoot 'NtnxPcInventory'
. (Join-Path $moduleRoot 'Common.ps1')
. (Join-Path $moduleRoot 'NtnxRest.ps1')
. (Join-Path $moduleRoot 'Collector.ps1')
. (Join-Path $moduleRoot 'ExcelExport.ps1')

$script:NtnxVerbose = $PSBoundParameters.ContainsKey('Verbose')

function Assert-ImportExcel {
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction Stop
        return
    }

    throw @'
The ImportExcel module is required to write the workbook.

  Online : Install-Module ImportExcel -Scope CurrentUser
  Offline: on a machine with internet access run
             Save-Module ImportExcel -Path .\modules
           copy the .\modules\ImportExcel folder to this host, then
             Import-Module .\modules\ImportExcel\<version>\ImportExcel.psd1
'@
}

function Read-NonEmpty {
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $value = (Read-Host -Prompt $Prompt).Trim()
        if ($value) { return $value }
        Write-Host 'Input cannot be empty.' -ForegroundColor Yellow
    }
}

function Read-YesNo {
    param([Parameter(Mandatory)][string]$Prompt)
    while ($true) {
        $value = (Read-Host -Prompt "$Prompt (y/n)").Trim().ToLower()
        if ($value -eq 'y') { return $true }
        if ($value -eq 'n') { return $false }
    }
}

function ConvertTo-ClusterList {
    param($Value)
    if (-not $Value) { return @('ALL') }
    $items = @($Value) | ForEach-Object { ([string]$_).Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ }
    if (-not $items) { return @('ALL') }
    return @($items)
}

function Get-InventoryTaskFromInteractive {
    Write-Host ('=' * 52)
    Write-Host '  Nutanix Prism Central - PC inventory (Excel)'
    Write-Host ('=' * 52)
    Write-Host ''
    Write-Host '  1 - Interactive (enter Prism Central details)'
    Write-Host '  2 - Credentials file (CSV or Excel)'

    $choice = ''
    while ($choice -notin @('1', '2')) { $choice = (Read-Host 'Choice (1 or 2)').Trim() }

    $tasks = [System.Collections.Generic.List[object]]::new()

    if ($choice -eq '2') {
        $path = Read-NonEmpty 'Path to credentials CSV or Excel'
        $clusterInput = (Read-Host 'Cluster names for all rows (comma-separated, Enter = ALL)').Trim()
        foreach ($t in (Get-InventoryTaskFromFile -Path $path -DefaultClusters (ConvertTo-ClusterList $clusterInput))) {
            $tasks.Add($t)
        }
        return $tasks
    }

    while ($true) {
        Write-Host ''
        Write-Host "--- PC $($tasks.Count + 1) ---"
        $ip = Read-NonEmpty 'Prism Central IP or FQDN'

        Write-Host '  Auth: 1 = Basic  2 = API key'
        $mode = ''
        while ($mode -notin @('1', '2')) { $mode = (Read-Host '  Select (1/2)').Trim() }

        if ($mode -eq '1') {
            $user = Read-NonEmpty '  Username'
            $secure = Read-Host '  Password' -AsSecureString
            $secret = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
                [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure))
            $auth = 'basic'
        }
        else {
            $user = Read-NonEmpty '  API key'
            $secret = ''
            $auth = 'api_key'
        }

        $clusterInput = (Read-Host '  Cluster names (comma-separated, Enter = ALL)').Trim()

        $tasks.Add([PSCustomObject]@{
            PcIp           = $ip
            PcUser         = $user
            PcSecret       = $secret
            AuthMode       = $auth
            Clusters       = ConvertTo-ClusterList $clusterInput
            ApiPathVersion = 'auto'
        })

        if (-not (Read-YesNo 'Add another Prism Central?')) { break }
    }

    return $tasks
}

function Get-InventoryTaskFromFile {
    param(
        [Parameter(Mandatory)][string]$Path,
        [string[]]$DefaultClusters = @('ALL')
    )

    $resolved = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path
    $extension = [IO.Path]::GetExtension($resolved).ToLower()

    if ($extension -eq '.csv') {
        $records = @(Import-Csv -LiteralPath $resolved)
    }
    else {
        Assert-ImportExcel
        $records = @(Import-Excel -Path $resolved)
    }

    if (-not $records.Count) { throw "Credentials file is empty: $resolved" }

    $columns = $records[0].PSObject.Properties.Name
    foreach ($required in @('pc_ip', 'pc_user')) {
        if ($columns -notcontains $required) {
            throw "File must include columns: pc_ip, pc_user (optional: pc_secret, auth_mode, clusters, api_path_version)"
        }
    }

    $tasks = [System.Collections.Generic.List[object]]::new()
    foreach ($record in $records) {
        $ip = [string](Get-Prop $record 'pc_ip' '')
        $ip = $ip.Trim()
        if (-not $ip) { continue }

        $auth = ([string](Get-Prop $record 'auth_mode' 'basic')).Trim().ToLower()
        if ($auth -notin @('basic', 'api_key')) { $auth = 'basic' }

        $secret = ([string](Get-Prop $record 'pc_secret' '')).Trim()
        if ($auth -eq 'basic' -and -not $secret) {
            Write-NtnxLog "Skipping ${ip}: basic auth requires pc_secret in file." -Level WARN
            continue
        }

        $rowClusters = ([string](Get-Prop $record 'clusters' '')).Trim()
        $version = ([string](Get-Prop $record 'api_path_version' '')).Trim().ToLower()
        if (-not $version) { $version = 'auto' }

        $tasks.Add([PSCustomObject]@{
            PcIp           = $ip
            PcUser         = ([string](Get-Prop $record 'pc_user' '')).Trim()
            PcSecret       = $(if ($auth -eq 'basic') { $secret } else { '' })
            AuthMode       = $auth
            Clusters       = $(if ($rowClusters) { ConvertTo-ClusterList $rowClusters } else { $DefaultClusters })
            ApiPathVersion = $version
        })
    }

    return $tasks
}

function Get-InventoryTaskFromParameter {
    # Environment variables let the script run unattended with injected secrets.
    $ips = $PcIp
    if (-not $ips -and $env:PRISM_CENTRAL_VIP) { $ips = @($env:PRISM_CENTRAL_VIP) }

    $user = $PcUser
    if (-not $user -and $env:PRISM_CENTRAL_USERNAME) { $user = $env:PRISM_CENTRAL_USERNAME }

    $secret = $PcSecret
    if (-not $secret -and $env:PRISM_CENTRAL_PASSWORD) { $secret = $env:PRISM_CENTRAL_PASSWORD }

    if (-not $ips) { return @() }
    if (-not $user) { throw 'PcUser is required (or set PRISM_CENTRAL_USERNAME).' }
    if ($AuthMode -eq 'basic' -and -not $secret) {
        throw 'PcSecret is required for basic auth (or set PRISM_CENTRAL_PASSWORD).'
    }

    $expanded = @($ips) | ForEach-Object { $_.Split(',') } | ForEach-Object { $_.Trim() } | Where-Object { $_ }

    return @($expanded | ForEach-Object {
        [PSCustomObject]@{
            PcIp           = $_
            PcUser         = $user
            PcSecret       = $secret
            AuthMode       = $AuthMode
            Clusters       = ConvertTo-ClusterList $Clusters
            ApiPathVersion = $ApiPathVersion
        }
    })
}

# --- Main -------------------------------------------------------------------------

Assert-ImportExcel

$tasks = @()
if ($PSCmdlet.ParameterSetName -eq 'File') {
    $tasks = @(Get-InventoryTaskFromFile -Path $CredentialsFile -DefaultClusters (ConvertTo-ClusterList $Clusters))
}
else {
    $tasks = @(Get-InventoryTaskFromParameter)
    if (-not $tasks.Count -or $Interactive) {
        $tasks = @(Get-InventoryTaskFromInteractive)
    }
}

if (-not $tasks.Count) {
    Write-NtnxLog 'No tasks to run.' -Level WARN
    exit 0
}

if (-not (Test-Path -LiteralPath $OutputPath)) {
    $null = New-Item -ItemType Directory -Path $OutputPath -Force
}
$resolvedOutput = (Resolve-Path -LiteralPath $OutputPath).Path
Write-NtnxLog "Output directory: $resolvedOutput"

if ($PSVersionTable.PSEdition -ne 'Core') {
    Write-NtnxLog 'Running on Windows PowerShell - VM stats will be collected serially. PowerShell 7+ enables parallel collection.' -Level WARN
}

$failures = 0
$index = 0

foreach ($task in $tasks) {
    $index++
    Write-NtnxLog "[$index/$($tasks.Count)] Prism Central: $($task.PcIp)"
    try {
        $result = Get-NtnxPcInventory -PcIp $task.PcIp -PcUser $task.PcUser -PcSecret $task.PcSecret `
            -AuthMode $task.AuthMode -ClusterNames $task.Clusters `
            -UnderProvisionPercent $UnderProvisionPercent -OverProvisionPercent $OverProvisionPercent `
            -ApiPathVersion $task.ApiPathVersion

        $workbook = Export-NtnxInventoryWorkbook -Result $result -OutputDirectory $resolvedOutput `
            -UnderProvisionPercent $UnderProvisionPercent -OverProvisionPercent $OverProvisionPercent

        Write-NtnxLog "Saved workbook -> $workbook"
        Write-Output $workbook
    }
    catch {
        $failures++
        Write-NtnxLog "Failed for $($task.PcIp): $($_.Exception.Message)" -Level ERROR
        Write-NtnxLog $_.ScriptStackTrace -Level DEBUG
    }
}

Write-NtnxLog "Done. $($tasks.Count - $failures)/$($tasks.Count) succeeded."
if ($failures -gt 0) { exit 1 }
exit 0
