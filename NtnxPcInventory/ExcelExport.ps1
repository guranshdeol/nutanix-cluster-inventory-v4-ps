<#
    Builds the timestamped multi-sheet workbook via the ImportExcel module.

    Row shading and usage colouring are applied as range-level conditional
    formatting rather than per-cell fills. The rendered result matches the
    openpyxl original, but styling a 10k-row VM sheet stays a handful of
    operations instead of tens of thousands.
#>

Set-StrictMode -Version Latest

# Visual theme - same palette as the Python exporter.
$script:ThemePrimaryDeep = '111827'   # slate-900, title bar
$script:ThemePrimary     = '1F2937'   # slate-800, section headers
$script:ThemeAccent      = '6D28D9'   # violet-700, header rows
$script:ThemeAccentLight = 'EDE9FE'   # violet-100, alternating rows
$script:ThemeBorder      = 'C7C3D9'
$script:ThemeKpiBg       = 'F5F3FF'
$script:ThemeKpiLabel    = '374151'
$script:ThemeKpiValue    = '5B21B6'
$script:ThemeOk          = '16A34A'
$script:ThemeWarn        = 'D97706'
$script:ThemeCrit        = 'DC2626'
$script:FillUsageHigh    = 'FECACA'
$script:FillUsageLow     = 'FEF3C7'

$script:CpuColorColumn = 'CPU Usage Color'
$script:MemColorColumn = 'Memory Usage Color'
$script:CpuUsageColumn = 'vCPU usage %'
$script:MemUsageColumn = 'Memory usage %'

function ConvertTo-DrawingColor {
    param([Parameter(Mandatory)][string]$Hex)
    $clean = $Hex.TrimStart('#')
    return [System.Drawing.Color]::FromArgb(
        [Convert]::ToInt32($clean.Substring(0, 2), 16),
        [Convert]::ToInt32($clean.Substring(2, 2), 16),
        [Convert]::ToInt32($clean.Substring(4, 2), 16)
    )
}

function Get-ExcelColumnLetter {
    param([Parameter(Mandatory)][int]$Index)
    $letter = ''
    while ($Index -gt 0) {
        $mod = ($Index - 1) % 26
        $letter = [char](65 + $mod) + $letter
        $Index = [math]::Floor(($Index - $mod) / 26)
    }
    return $letter
}

function Get-SafeSheetName {
    param(
        [Parameter(Mandatory)][string]$Name,
        # Starts empty on the first sheet; a mandatory parameter rejects an
        # empty collection unless this is declared.
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.HashSet[string]]$Used,
        [int]$MaxLength = 31
    )

    $clean = ($Name -replace '[\[\]:\*\?/\\]', '_').Trim()
    if (-not $clean) { $clean = 'Sheet' }
    if ($clean.Length -gt $MaxLength) { $clean = $clean.Substring(0, $MaxLength) }

    if ($Used.Add($clean)) { return $clean }

    $i = 2
    while ($true) {
        $suffix = "_$i"
        $trimmed = $Name
        if ($trimmed.Length -gt ($MaxLength - $suffix.Length)) {
            $trimmed = $trimmed.Substring(0, $MaxLength - $suffix.Length)
        }
        $candidate = ($trimmed + $suffix) -replace '[\[\]:\*\?/\\]', '_'
        if ($Used.Add($candidate)) { return $candidate }
        $i++
    }
}

function ConvertTo-ExcelRow {
    <#
        Turn the ordered hashtables produced by the collector into objects that
        Export-Excel can consume, optionally dropping helper columns.
    #>
    param(
        $Rows,
        [string[]]$ExcludeColumns = @()
    )

    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $copy = [ordered]@{}
        foreach ($key in $row.Keys) {
            if ($ExcludeColumns -contains $key) { continue }
            $value = $row[$key]
            # Excel rejects some control characters and caps cell text length.
            if ($value -is [string]) {
                $value = $value -replace '[\x00-\x08\x0B\x0C\x0E-\x1F]', ''
                if ($value.Length -gt 32766) { $value = $value.Substring(0, 32766) }
            }
            $copy[$key] = $value
        }
        $out.Add([PSCustomObject]$copy)
    }
    return , $out.ToArray()
}

function Set-NtnxTableStyle {
    <# Header banding, borders, alternating rows, freeze panes and auto-filter. #>
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][int]$RowCount,
        [Parameter(Mandatory)][int]$ColumnCount
    )

    if ($ColumnCount -lt 1) { return }

    $lastCol = Get-ExcelColumnLetter $ColumnCount
    $headerRange = "A1:${lastCol}1"

    $header = $Worksheet.Cells[$headerRange]
    $header.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $header.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemeAccent))
    $header.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $header.Style.Font.Bold = $true
    $header.Style.Font.Size = 11
    $header.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $header.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $header.Style.WrapText = $true
    $Worksheet.Row(1).Height = 26

    if ($RowCount -gt 0) {
        $dataRange = "A1:${lastCol}$($RowCount + 1)"
        $all = $Worksheet.Cells[$dataRange]
        $borderColor = ConvertTo-DrawingColor $script:ThemeBorder
        $all.Style.Border.Top.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $all.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $all.Style.Border.Left.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $all.Style.Border.Right.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
        $all.Style.Border.Top.Color.SetColor($borderColor)
        $all.Style.Border.Bottom.Color.SetColor($borderColor)
        $all.Style.Border.Left.Color.SetColor($borderColor)
        $all.Style.Border.Right.Color.SetColor($borderColor)

        # Alternating row tint as one rule instead of per-cell fills.
        Add-ConditionalFormatting -Worksheet $Worksheet -Address "A2:${lastCol}$($RowCount + 1)" `
            -RuleType Expression -ConditionValue 'MOD(ROW(),2)=1' `
            -BackgroundColor (ConvertTo-DrawingColor $script:ThemeAccentLight)

        $Worksheet.Cells["A1:${lastCol}$($RowCount + 1)"].AutoFilter = $true
    }

    $Worksheet.View.FreezePanes(2, 1)
}

function Add-NtnxUsageFormatting {
    <#
        Reproduces the CPU/Memory usage shading. The Python version precomputed a
        helper colour column; the same thresholds expressed as a formula rule let
        us drop those helper columns from the sheet entirely.
    #>
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$Headers,
        [Parameter(Mandatory)][int]$RowCount,
        [double]$UnderProvisionPercent,
        [double]$OverProvisionPercent
    )

    if ($RowCount -lt 1) { return }

    foreach ($columnName in @($script:CpuUsageColumn, $script:MemUsageColumn)) {
        $index = [array]::IndexOf($Headers, $columnName)
        if ($index -lt 0) { continue }

        $letter = Get-ExcelColumnLetter ($index + 1)
        $range = "${letter}2:${letter}$($RowCount + 1)"

        Add-ConditionalFormatting -Worksheet $Worksheet -Address $range -RuleType Expression `
            -ConditionValue "AND(ISNUMBER(${letter}2),${letter}2>$UnderProvisionPercent)" `
            -BackgroundColor (ConvertTo-DrawingColor $script:FillUsageHigh)

        Add-ConditionalFormatting -Worksheet $Worksheet -Address $range -RuleType Expression `
            -ConditionValue "AND(ISNUMBER(${letter}2),${letter}2<$OverProvisionPercent)" `
            -BackgroundColor (ConvertTo-DrawingColor $script:FillUsageLow)
    }
}

function Add-NtnxSeverityFormatting {
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$Headers,
        [Parameter(Mandatory)][int]$RowCount
    )

    if ($RowCount -lt 1) { return }
    $index = [array]::IndexOf($Headers, 'Severity')
    if ($index -lt 0) { return }

    $letter = Get-ExcelColumnLetter ($index + 1)
    $range = "${letter}2:${letter}$($RowCount + 1)"

    Add-ConditionalFormatting -Worksheet $Worksheet -Address $range -RuleType Expression `
        -ConditionValue "EXACT(${letter}2,""CRITICAL"")" `
        -BackgroundColor (ConvertTo-DrawingColor $script:FillUsageHigh) `
        -ForegroundColor (ConvertTo-DrawingColor $script:ThemeCrit) -Bold

    Add-ConditionalFormatting -Worksheet $Worksheet -Address $range -RuleType Expression `
        -ConditionValue "EXACT(${letter}2,""WARNING"")" `
        -BackgroundColor (ConvertTo-DrawingColor $script:FillUsageLow) `
        -ForegroundColor (ConvertTo-DrawingColor $script:ThemeWarn) -Bold
}

function Add-NtnxDataSheet {
    <# Write one data sheet and return the worksheet plus its header list. #>
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)][string]$Name,
        $Rows,
        [string[]]$ExcludeColumns = @()
    )

    $objects = ConvertTo-ExcelRow -Rows $Rows -ExcludeColumns $ExcludeColumns

    if (-not $objects.Count) {
        $ws = $Package.Workbook.Worksheets.Add($Name)
        $ws.Cells['A1'].Value = 'No data available.'
        $ws.Cells['A1'].Style.Font.Italic = $true
        return @{ Worksheet = $ws; Headers = @(); RowCount = 0 }
    }

    $null = $objects | Export-Excel -ExcelPackage $Package -WorksheetName $Name -AutoSize -PassThru
    $ws = $Package.Workbook.Worksheets[$Name]

    $headers = @($objects[0].PSObject.Properties.Name)
    Set-NtnxTableStyle -Worksheet $ws -RowCount $objects.Count -ColumnCount $headers.Count

    return @{ Worksheet = $ws; Headers = $headers; RowCount = $objects.Count }
}

function Add-NtnxKpiTile {
    <# 3-column x 3-row KPI tile anchored at (Top, Left). #>
    param(
        [Parameter(Mandatory)]$Worksheet,
        [Parameter(Mandatory)][int]$Top,
        [Parameter(Mandatory)][int]$Left,
        [Parameter(Mandatory)][string]$Label,
        $Value,
        [string]$ValueColor
    )

    if (-not $ValueColor) { $ValueColor = $script:ThemeKpiValue }
    $bottom = $Top + 2
    $right = $Left + 2
    $leftLetter = Get-ExcelColumnLetter $Left
    $rightLetter = Get-ExcelColumnLetter $right

    $tile = $Worksheet.Cells["${leftLetter}${Top}:${rightLetter}${bottom}"]
    $tile.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $tile.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemeKpiBg))
    $tile.Style.Border.BorderAround(
        [OfficeOpenXml.Style.ExcelBorderStyle]::Medium,
        (ConvertTo-DrawingColor $script:ThemeAccent))

    $labelRange = $Worksheet.Cells["${leftLetter}${Top}:${rightLetter}${Top}"]
    $labelRange.Merge = $true
    $Worksheet.Cells["${leftLetter}${Top}"].Value = $Label
    $labelRange.Style.Font.Bold = $true
    $labelRange.Style.Font.Size = 10
    $labelRange.Style.Font.Color.SetColor((ConvertTo-DrawingColor $script:ThemeKpiLabel))
    $labelRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $labelRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    $valueRange = $Worksheet.Cells["${leftLetter}$($Top + 1):${rightLetter}${bottom}"]
    $valueRange.Merge = $true
    $Worksheet.Cells["${leftLetter}$($Top + 1)"].Value = $Value
    $valueRange.Style.Font.Bold = $true
    $valueRange.Style.Font.Size = 18
    $valueRange.Style.Font.Color.SetColor((ConvertTo-DrawingColor $ValueColor))
    $valueRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $valueRange.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center

    $Worksheet.Row($Top).Height = 22
    $Worksheet.Row($Top + 1).Height = 26
    $Worksheet.Row($bottom).Height = 26
}

function Add-NtnxSummarySheet {
    param(
        [Parameter(Mandatory)]$Package,
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$Name
    )

    $ws = $Package.Workbook.Worksheets.Add($Name)
    $summary = $Result.PcSummary

    # Banner
    $banner = $ws.Cells['A1:L2']
    $banner.Merge = $true
    $ws.Cells['A1'].Value = "Nutanix Prism Central Inventory - $($Result.PcName)"
    $banner.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $banner.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemePrimaryDeep))
    $banner.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $banner.Style.Font.Bold = $true
    $banner.Style.Font.Size = 20
    $banner.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $banner.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $ws.Row(1).Height = 30
    $ws.Row(2).Height = 22

    $clusterList = ($Result.ClustersProcessed -join ', ')
    if (-not $clusterList) { $clusterList = 'None' }

    $sub = $ws.Cells['A3:L3']
    $sub.Merge = $true
    $ws.Cells['A3'].Value = "PC version: $($summary.pc_version)    -    Endpoint: $($Result.PcIp)    -    " +
        "Generated (UTC): $($Result.GeneratedAtUtc.ToString('yyyy-MM-dd HH:mm:ss'))    -    Clusters: $clusterList"
    $sub.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $sub.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemePrimaryDeep))
    $sub.Style.Font.Color.SetColor((ConvertTo-DrawingColor 'E5E7EB'))
    $sub.Style.Font.Italic = $true
    $sub.Style.Font.Size = 11
    $sub.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $sub.Style.VerticalAlignment = [OfficeOpenXml.Style.ExcelVerticalAlignment]::Center
    $ws.Row(3).Height = 22

    $critical = $summary.critical_alerts
    $warning = $summary.warning_alerts
    $criticalColor = if ($critical -gt 0) { $script:ThemeCrit } else { $script:ThemeOk }
    $warningColor = if ($warning -gt 0) { $script:ThemeWarn } else { $script:ThemeOk }

    $tiles = @(
        @{ Label = 'Clusters';            Value = $summary.cluster_count }
        @{ Label = 'Hypervisor Hosts';    Value = $summary.host_count }
        @{ Label = 'Virtual Machines';    Value = $summary.vm_count }
        @{ Label = 'Storage Containers';  Value = $summary.storage_container_count }
        @{ Label = 'Subnets';             Value = $summary.subnet_count }
        @{ Label = 'Virtual Switches';    Value = $summary.virtual_switch_count }
        @{ Label = 'Recovery Points';     Value = $summary.recovery_point_count }
        @{ Label = 'Protection Policies'; Value = $summary.protection_policy_count }
        @{ Label = 'Host Disks';          Value = $summary.disk_count }
        @{ Label = 'Licenses';            Value = $summary.license_count }
        @{ Label = 'Critical Alerts';     Value = $critical; Color = $criticalColor }
        @{ Label = 'Warning Alerts';      Value = $warning;  Color = $warningColor }
    )

    $startRow = 5
    for ($i = 0; $i -lt $tiles.Count; $i++) {
        $tile = $tiles[$i]
        $left = 1 + (($i % 4) * 3)
        $top = $startRow + ([math]::Floor($i / 4) * 4)
        $color = if ($tile.ContainsKey('Color')) { $tile.Color } else { $null }
        Add-NtnxKpiTile -Worksheet $ws -Top $top -Left $left -Label $tile.Label -Value $tile.Value -ValueColor $color
    }

    $tileRows = [math]::Ceiling($tiles.Count / 4)
    $sectionRow = $startRow + ($tileRows * 4) + 1

    $section = $ws.Cells["A${sectionRow}:L${sectionRow}"]
    $section.Merge = $true
    $ws.Cells["A${sectionRow}"].Value = 'Cluster snapshot'
    $section.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $section.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemePrimary))
    $section.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $section.Style.Font.Bold = $true
    $section.Style.Font.Size = 12
    $section.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    $section.Style.Indent = 1
    $ws.Row($sectionRow).Height = 24

    $headers = @('Cluster', 'AOS', 'Hypervisor', 'Nodes', 'VMs', 'Cluster VIP',
                 'FT (current)', 'Domain', 'Pulse', 'Critical', 'Warning', 'Open')
    $headerRow = $sectionRow + 1
    # The comma in a two-argument indexer binds tighter than '+', so the column
    # index has to be computed first rather than inline.
    for ($c = 0; $c -lt $headers.Count; $c++) {
        $column = $c + 1
        $ws.Cells[$headerRow, $column].Value = $headers[$c]
    }
    $headerRange = $ws.Cells["A${headerRow}:L${headerRow}"]
    $headerRange.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $headerRange.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:ThemeAccent))
    $headerRange.Style.Font.Color.SetColor([System.Drawing.Color]::White)
    $headerRange.Style.Font.Bold = $true
    $headerRange.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Center
    $ws.Row($headerRow).Height = 24

    $fields = @('Cluster Name', 'AOS Version', 'Hypervisor', 'Number of Nodes', 'VM Count',
                'Cluster Virtual IP', 'Cluster FT (Current)', 'Domain Awareness', 'Pulse',
                'Critical Alerts', 'Warning Alerts', 'Total Open Alerts')

    $r = $headerRow
    foreach ($row in @($Result.ClusterOverviewRows)) {
        $r++
        for ($c = 0; $c -lt $fields.Count; $c++) {
            $column = $c + 1
            $ws.Cells[$r, $column].Value = $row[$fields[$c]]
        }

        $criticalCount = $row['Critical Alerts']
        if ($criticalCount -is [int] -and $criticalCount -gt 0) {
            $cell = $ws.Cells[$r, 10]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:FillUsageHigh))
            $cell.Style.Font.Bold = $true
            $cell.Style.Font.Color.SetColor((ConvertTo-DrawingColor $script:ThemeCrit))
        }
        $warningCount = $row['Warning Alerts']
        if ($warningCount -is [int] -and $warningCount -gt 0) {
            $cell = $ws.Cells[$r, 11]
            $cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $cell.Style.Fill.BackgroundColor.SetColor((ConvertTo-DrawingColor $script:FillUsageLow))
            $cell.Style.Font.Bold = $true
            $cell.Style.Font.Color.SetColor((ConvertTo-DrawingColor $script:ThemeWarn))
        }
    }

    if ($r -gt $headerRow) {
        $body = $ws.Cells["A$($headerRow):L${r}"]
        $borderColor = ConvertTo-DrawingColor $script:ThemeBorder
        foreach ($side in @('Top', 'Bottom', 'Left', 'Right')) {
            $body.Style.Border.$side.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $body.Style.Border.$side.Color.SetColor($borderColor)
        }
    }

    $widths = @(22, 14, 22, 9, 9, 18, 14, 16, 12, 12, 12, 12)
    for ($i = 0; $i -lt $widths.Count; $i++) {
        $ws.Column($i + 1).Width = $widths[$i]
    }

    return $ws
}

function Export-NtnxInventoryWorkbook {
    <#
    .SYNOPSIS
        Write one .xlsx for a Prism Central inventory result.
    .OUTPUTS
        The full path of the workbook.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]$Result,
        [Parameter(Mandatory)][string]$OutputDirectory,
        [double]$UnderProvisionPercent = 75.0,
        [double]$OverProvisionPercent = 30.0
    )

    if (-not (Test-Path -LiteralPath $OutputDirectory)) {
        $null = New-Item -ItemType Directory -Path $OutputDirectory -Force
    }

    $safeName = ($Result.PcName -replace '\s', '_') -replace '[\[\]:\*\?/\\]', '_'
    if ($safeName.Length -gt 40) { $safeName = $safeName.Substring(0, 40) }
    $fileName = "Nutanix_PC_Inventory_${safeName}_$($Result.TimestampSlug).xlsx"
    $path = Join-Path -Path (Resolve-Path -LiteralPath $OutputDirectory).Path -ChildPath $fileName

    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }

    $package = Open-ExcelPackage -Path $path -Create
    $used = [System.Collections.Generic.HashSet[string]]::new()

    try {
        $null = Add-NtnxSummarySheet -Package $package -Result $Result -Name (Get-SafeSheetName 'Summary' $used)

        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Cluster Overview' $used) -Rows $Result.ClusterOverviewRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Hypervisor Hosts' $used) -Rows $Result.HostRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Host Disks' $used) -Rows $Result.HostDiskRows

        $vmSheet = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Virtual Machines' $used) `
            -Rows $Result.VmRows -ExcludeColumns @($script:CpuColorColumn, $script:MemColorColumn)
        Add-NtnxUsageFormatting -Worksheet $vmSheet.Worksheet -Headers $vmSheet.Headers -RowCount $vmSheet.RowCount `
            -UnderProvisionPercent $UnderProvisionPercent -OverProvisionPercent $OverProvisionPercent

        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Storage Containers' $used) -Rows $Result.StorageRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Networks' $used) -Rows $Result.SubnetRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Virtual Switches' $used) -Rows $Result.VSwitchRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Recovery Points' $used) -Rows $Result.RecoveryPointRows
        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Protection Policies' $used) -Rows $Result.ProtectionPolicyRows

        $alertSheet = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Alerts (Critical & Warning)' $used) -Rows $Result.AlertRows
        Add-NtnxSeverityFormatting -Worksheet $alertSheet.Worksheet -Headers $alertSheet.Headers -RowCount $alertSheet.RowCount

        $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName 'Licenses' $used) -Rows $Result.LicenseRows

        # Per-cluster breakouts only make sense with more than one cluster.
        if (@($Result.ClustersProcessed).Count -gt 1) {
            foreach ($clusterName in $Result.ClustersProcessed) {
                $vmSubset = @($Result.VmRows | Where-Object { $_['Parent Cluster'] -eq $clusterName })
                $hostSubset = @($Result.HostRows | Where-Object { $_['Cluster Name'] -eq $clusterName })

                if ($vmSubset.Count) {
                    $sheet = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName "VMs_$clusterName" $used) `
                        -Rows $vmSubset -ExcludeColumns @($script:CpuColorColumn, $script:MemColorColumn)
                    Add-NtnxUsageFormatting -Worksheet $sheet.Worksheet -Headers $sheet.Headers -RowCount $sheet.RowCount `
                        -UnderProvisionPercent $UnderProvisionPercent -OverProvisionPercent $OverProvisionPercent
                }
                if ($hostSubset.Count) {
                    $null = Add-NtnxDataSheet -Package $package -Name (Get-SafeSheetName "Hosts_$clusterName" $used) -Rows $hostSubset
                }
            }
        }

        Close-ExcelPackage $package
    }
    catch {
        if ($package) { Close-ExcelPackage $package -NoSave }
        throw
    }

    return $path
}
