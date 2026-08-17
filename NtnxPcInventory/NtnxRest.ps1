<#
    REST layer for the Prism Central v4 API.

    Replaces the eight Nutanix Python SDKs with direct REST calls, which also
    removes the need for the SDK/PC version-negotiation workaround: instead of
    rewriting SDK request paths, we probe each namespace once and pin the minor
    version the PC actually exposes.
#>

Set-StrictMode -Version Latest

# Namespaces used by the inventory, with the candidate minors to probe.
$script:NtnxNamespaceProbe = [ordered]@{
    clustermgmt    = 'clustermgmt/{0}/config/clusters'
    vmm            = 'vmm/{0}/ahv/config/vms'
    prism          = 'prism/{0}/config/categories'
    networking     = 'networking/{0}/config/subnets'
    dataprotection = 'dataprotection/{0}/config/recovery-points'
    monitoring     = 'monitoring/{0}/serviceability/alerts'
    licensing      = 'licensing/{0}/config/licenses'
    datapolicies   = 'datapolicies/{0}/config/protection-policies'
}

$script:NtnxCandidateVersions = @('v4.0', 'v4.1', 'v4.2', 'v4.3')

function New-NtnxSession {
    <#
    .SYNOPSIS
        Build a connection context for one Prism Central.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$PcIp,
        [Parameter(Mandatory)][string]$User,
        [string]$Secret,
        [ValidateSet('basic', 'api_key')][string]$AuthMode = 'basic',
        [int]$Port = 9440,
        [string]$ApiPathVersion = 'auto',
        [int]$TimeoutSec = 120
    )

    # Accept "host" or "host:port".
    $pcHost = $PcIp
    if ($PcIp -match '^(?<h>[^:]+):(?<p>\d+)$') {
        $pcHost = $Matches['h']
        $Port = [int]$Matches['p']
    }

    if ($AuthMode -eq 'basic' -and [string]::IsNullOrWhiteSpace($Secret)) {
        throw "Secret is required for basic auth."
    }

    $headers = @{ Accept = 'application/json' }
    if ($AuthMode -eq 'api_key') {
        $headers['X-ntnx-api-key'] = $User
    }
    else {
        # Prism Central does not issue a 401 challenge, so send the header up front.
        $pair = '{0}:{1}' -f $User, $Secret
        $headers['Authorization'] = 'Basic ' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($pair))
    }

    $session = [PSCustomObject]@{
        Host        = $pcHost
        Port        = $Port
        BaseUrl     = "https://${pcHost}:${Port}/api"
        Headers     = $headers
        TimeoutSec  = $TimeoutSec
        IsCore      = ($PSVersionTable.PSEdition -eq 'Core')
        Versions    = @{}
        PinnedVersion = $ApiPathVersion
        CategoryCache = @{}
    }

    Enable-NtnxCertificateBypass -Session $session
    Resolve-NtnxApiVersions -Session $session
    return $session
}

function Enable-NtnxCertificateBypass {
    <#
        Prism Central ships a self-signed certificate. PowerShell 7 handles this
        per-request with -SkipCertificateCheck; Windows PowerShell 5.1 needs a
        process-wide callback and an explicit TLS 1.2.
    #>
    param([Parameter(Mandatory)]$Session)

    if ($Session.IsCore) { return }

    [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
    if ($null -eq [System.Net.ServicePointManager]::ServerCertificateValidationCallback) {
        [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
    }
}

function Invoke-NtnxApi {
    <#
    .SYNOPSIS
        Single authenticated GET against the PC API, with retry on throttling.
    .PARAMETER Path
        Path below /api, e.g. 'clustermgmt/v4.0/config/clusters'.
    .PARAMETER Query
        Ordered hashtable of query parameters. OData keys keep their '$' prefix.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Path,
        [hashtable]$Query,
        [hashtable]$ExtraHeaders,
        [int]$MaxRetries = 4,
        [switch]$Quiet
    )

    $url = "$($Session.BaseUrl)/$Path"
    if ($Query -and $Query.Count) {
        $parts = foreach ($k in $Query.Keys) {
            if ($null -eq $Query[$k] -or $Query[$k] -eq '') { continue }
            '{0}={1}' -f [uri]::EscapeDataString($k), [uri]::EscapeDataString([string]$Query[$k])
        }
        if ($parts) { $url = $url + '?' + ($parts -join '&') }
    }

    $headers = @{}
    foreach ($k in $Session.Headers.Keys) { $headers[$k] = $Session.Headers[$k] }
    if ($ExtraHeaders) { foreach ($k in $ExtraHeaders.Keys) { $headers[$k] = $ExtraHeaders[$k] } }

    $params = @{
        Method      = 'Get'
        Uri         = $url
        Headers     = $headers
        TimeoutSec  = $Session.TimeoutSec
        ErrorAction = 'Stop'
    }
    if ($Session.IsCore) { $params['SkipCertificateCheck'] = $true }

    $attempt = 0
    while ($true) {
        try {
            return Invoke-RestMethod @params
        }
        catch {
            $status = $null
            if ($_.Exception.PSObject.Properties.Name -contains 'Response' -and $_.Exception.Response) {
                try { $status = [int]$_.Exception.Response.StatusCode } catch { $status = $null }
            }

            # 429/503 are the PC rate limiter; back off and retry.
            if ($attempt -lt $MaxRetries -and ($status -eq 429 -or $status -eq 503)) {
                $attempt++
                $delay = [math]::Min(30, [math]::Pow(2, $attempt))
                Write-NtnxLog "Throttled (HTTP $status) on $Path - retrying in ${delay}s" -Level DEBUG
                Start-Sleep -Seconds $delay
                continue
            }

            if (-not $Quiet) {
                Write-NtnxLog "GET $Path failed: $($_.Exception.Message)" -Level WARN
            }
            throw
        }
    }
}

function Get-NtnxList {
    <#
    .SYNOPSIS
        Page through a v4 list endpoint and return every entity.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Path,
        [string]$Filter,
        [string]$OrderBy,
        [hashtable]$ExtraHeaders,
        [int]$Limit = 100,
        [int]$MaxRows = 0
    )

    $all = [System.Collections.Generic.List[object]]::new()
    $page = 0

    while ($true) {
        $query = @{ '$page' = $page; '$limit' = $Limit }
        if ($Filter) { $query['$filter'] = $Filter }
        if ($OrderBy) { $query['$orderby'] = $OrderBy }

        try {
            $resp = Invoke-NtnxApi -Session $Session -Path $Path -Query $query -ExtraHeaders $ExtraHeaders
        }
        catch {
            break
        }

        $data = $null
        if ($resp -and $resp.PSObject.Properties.Name -contains 'data') { $data = $resp.data }
        if (-not $data) { break }

        # A single-object response still needs to become a list.
        $batch = @($data)
        foreach ($item in $batch) {
            $all.Add($item)
            if ($MaxRows -gt 0 -and $all.Count -ge $MaxRows) { break }
        }

        if ($MaxRows -gt 0 -and $all.Count -ge $MaxRows) { break }
        if ($batch.Count -lt $Limit) { break }
        $page++
    }

    return $all
}

function Resolve-NtnxApiVersions {
    <#
    .SYNOPSIS
        Determine which v4 minor each namespace exposes on this PC.
    .DESCRIPTION
        Newer SDKs/clients assume a higher minor than older PC builds serve
        (pc.7.5.0.x exposes only v4.0 for some namespaces), which surfaces as
        404s. Probing once up front and pinning the result avoids that entirely.
        Namespaces that answer on none of the candidates are marked unavailable
        and their sheets degrade to empty, matching the Python behaviour.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)]$Session)

    foreach ($ns in $script:NtnxNamespaceProbe.Keys) {
        $template = $script:NtnxNamespaceProbe[$ns]

        if ($Session.PinnedVersion -and $Session.PinnedVersion -ne 'auto') {
            $Session.Versions[$ns] = $Session.PinnedVersion
            continue
        }

        $found = $null
        foreach ($v in $script:NtnxCandidateVersions) {
            $path = $template -f $v
            try {
                Invoke-NtnxApi -Session $Session -Path $path -Query @{ '$limit' = 1 } -MaxRetries 1 -Quiet | Out-Null
                $found = $v
                break
            }
            catch {
                continue
            }
        }

        $Session.Versions[$ns] = $found
        if ($found) {
            Write-NtnxLog "Namespace '$ns' -> $found" -Level DEBUG
        }
        else {
            Write-NtnxLog "Namespace '$ns' is unavailable on this Prism Central; related sheets will be empty." -Level WARN
        }
    }
}

function Get-NtnxPath {
    <#
    .SYNOPSIS
        Build a namespace path using the minor version resolved for this PC.
        Returns $null when the namespace is unavailable.
    #>
    param(
        [Parameter(Mandatory)]$Session,
        [Parameter(Mandatory)][string]$Namespace,
        [Parameter(Mandatory)][string]$Suffix
    )

    if (-not $Session.Versions.ContainsKey($Namespace)) { return $null }
    $v = $Session.Versions[$Namespace]
    if (-not $v) { return $null }
    return "$Namespace/$v/$Suffix"
}
