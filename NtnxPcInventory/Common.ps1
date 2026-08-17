<#
    Shared helpers: logging, unit conversion, and safe access to the deeply
    nested optional fields the v4 API returns.
#>

Set-StrictMode -Version Latest

$script:NtnxGiB = [math]::Pow(1024, 3)

# Set before the entry point can override it; StrictMode rejects reads of an
# undefined variable.
$script:NtnxVerbose = $false

function Write-NtnxLog {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, Position = 0)][string]$Message,
        [ValidateSet('DEBUG', 'INFO', 'WARN', 'ERROR')][string]$Level = 'INFO'
    )

    if ($Level -eq 'DEBUG' -and -not $script:NtnxVerbose) { return }

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    switch ($Level) {
        'ERROR' { Write-Host $line -ForegroundColor Red }
        'WARN'  { Write-Host $line -ForegroundColor Yellow }
        'DEBUG' { Write-Host $line -ForegroundColor DarkGray }
        default { Write-Host $line }
    }
}

function Get-Prop {
    <#
    .SYNOPSIS
        Safely read a dotted property path from a JSON object.
    .DESCRIPTION
        Returns $Default when any link in the chain is missing or null. Needed
        because Set-StrictMode makes a missing property a terminating error, and
        most v4 response fields are optional.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Position = 0)]$Object,
        [Parameter(Mandatory, Position = 1)][string]$Path,
        [Parameter(Position = 2)]$Default = $null
    )

    $current = $Object
    foreach ($segment in $Path.Split('.')) {
        if ($null -eq $current) { return $Default }
        if ($current -is [System.Collections.IDictionary]) {
            if (-not $current.Contains($segment)) { return $Default }
            $current = $current[$segment]
            continue
        }
        $prop = $current.PSObject.Properties[$segment]
        if (-not $prop) { return $Default }
        $current = $prop.Value
    }

    if ($null -eq $current) { return $Default }
    return $current
}

function Get-PropText {
    <# Same as Get-Prop but always returns a non-empty string (default 'NA'). #>
    param(
        [Parameter(Position = 0)]$Object,
        [Parameter(Mandatory, Position = 1)][string]$Path,
        [Parameter(Position = 2)][string]$Default = 'NA'
    )

    $value = Get-Prop $Object $Path $null
    if ($null -eq $value) { return $Default }
    if ($value -is [array]) {
        $joined = ($value | ForEach-Object { [string]$_ }) -join ', '
        if ([string]::IsNullOrWhiteSpace($joined)) { return $Default }
        return $joined
    }
    $text = [string]$value
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
    return $text
}

function Get-PropBoolText {
    <# Booleans are rendered as Yes/No throughout the workbook. #>
    param(
        [Parameter(Position = 0)]$Object,
        [Parameter(Mandatory, Position = 1)][string]$Path,
        [string]$True_ = 'Yes',
        [string]$False_ = 'No'
    )

    if (Get-Prop $Object $Path $false) { return $True_ }
    return $False_
}

function Get-NtnxIpValue {
    <#
        IP fields arrive in several shapes: a plain string, {value=...}, or
        {ipv4={value=...}} / {ipv6={value=...}}.
    #>
    param($IpObject)

    if ($null -eq $IpObject) { return 'NA' }
    if ($IpObject -is [string]) { return $IpObject }

    foreach ($path in @('value', 'ipv4.value', 'ipv6.value')) {
        $v = Get-Prop $IpObject $path $null
        if ($v) { return [string]$v }
    }
    return 'NA'
}

function ConvertTo-Gib {
    param(
        $Bytes,
        [int]$Digits = 2,
        [switch]$Floor
    )

    if (-not $Bytes) { return 0 }
    $gib = [double]$Bytes / $script:NtnxGiB
    if ($Floor) { return [math]::Floor($gib) }
    return [math]::Round($gib, $Digits)
}

function Get-NtnxAverage {
    <#
        v4 stat series come back as [{timestamp, value}, ...]. The Python version
        averages the whole down-sampled window, so we do the same, and return 0
        when the series is empty or non-numeric.
    #>
    param($Series)

    if (-not $Series) { return 0 }

    $values = foreach ($item in @($Series)) {
        if ($null -eq $item) { continue }
        if ($item -is [int] -or $item -is [long] -or $item -is [double]) { $item; continue }
        $v = Get-Prop $item 'value' $null
        if ($null -ne $v) { $v }
    }

    $values = @($values)
    if (-not $values.Count) { return 0 }
    return [math]::Round(($values | Measure-Object -Average).Average)
}

function Get-NtnxUtcWindow {
    <# Last 24 hours, matching the Python collector. #>
    $end = [DateTime]::UtcNow
    $start = $end.AddHours(-24)
    return @{
        Start = $start.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
        End   = $end.ToString('yyyy-MM-ddTHH:mm:ss.fffZ')
    }
}
