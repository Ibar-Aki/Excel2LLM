[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$ExcelPath,
    [string]$OutputDir,
    [switch]$CollectStyles,
    [switch]$SkipStyles,
    [switch]$NoRecalculate
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $OutputDir) {
    $OutputDir = Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'output'
}

if ($SkipStyles) {
    $CollectStyles = $false
}

$warnings = [System.Collections.Generic.List[string]]::new()
$excel = $null
$workbook = $null
$usedRange = $null

try {
    $resolvedExcelPath = Resolve-AbsolutePath -Path $ExcelPath
    Ensure-Directory -Path $OutputDir

    $workbookJsonPath = Join-Path $OutputDir 'workbook.json'
    $stylesJsonPath = Join-Path $OutputDir 'styles.json'
    $manifestJsonPath = Join-Path $OutputDir 'manifest.json'

    $excel = New-ExcelApplication
    $workbook = $excel.Workbooks.Open($resolvedExcelPath, 0, $true)
    if (-not $NoRecalculate) {
        try {
            $excel.CalculateFullRebuild()
        }
        catch {
            Add-WarningMessage -Warnings $warnings -Message "Recalculation failed: $($_.Exception.Message)"
        }
    }

    $sheets = New-Object System.Collections.Generic.List[object]
    $cells = New-Object System.Collections.Generic.List[object]
    $styles = New-Object System.Collections.Generic.List[object]
    $mergedRanges = New-Object System.Collections.Generic.List[object]
    $globalMergedKeys = [System.Collections.Generic.HashSet[string]]::new()

    $totalFormulaCount = 0
    $totalCellCount = 0

    foreach ($sheet in $workbook.Worksheets) {
        $usedRange = $sheet.UsedRange
        $rangeInfo = Get-UsedRangeInfo -UsedRange $usedRange
        $freezePanes = Get-WorksheetFreezeState -Excel $excel -Worksheet $sheet

        $sheetMergedRanges = New-Object System.Collections.Generic.List[object]
        $sheetMergedKeys = [System.Collections.Generic.HashSet[string]]::new()
        $hiddenRows = New-Object System.Collections.Generic.List[int]
        $hiddenColumns = New-Object System.Collections.Generic.List[string]
        $rowHeights = New-Object System.Collections.Generic.List[object]
        $columnWidths = New-Object System.Collections.Generic.List[object]

        for ($rowIndex = $rangeInfo.first_row; $rowIndex -le $rangeInfo.last_row; $rowIndex++) {
            $rowRange = $null
            try {
                $rowRange = $sheet.Rows.Item($rowIndex)
                if ($rowRange.Hidden) {
                    $hiddenRows.Add($rowIndex)
                }
                $rowHeights.Add([ordered]@{
                    row = $rowIndex
                    height = [double]$rowRange.RowHeight
                })
            }
            finally {
                if ($null -ne $rowRange) {
                    Release-ComReference $rowRange
                }
            }
        }

        for ($columnIndex = $rangeInfo.first_column; $columnIndex -le $rangeInfo.last_column; $columnIndex++) {
            $columnRange = $null
            try {
                $columnRange = $sheet.Columns.Item($columnIndex)
                $columnLetters = (Convert-CoordinateToA1 -Row 1 -Column $columnIndex) -replace '\d', ''
                if ($columnRange.Hidden) {
                    $hiddenColumns.Add($columnLetters)
                }
                $columnWidths.Add([ordered]@{
                    column = $columnLetters
                    width = [double]$columnRange.ColumnWidth
                })
            }
            finally {
                if ($null -ne $columnRange) {
                    Release-ComReference $columnRange
                }
            }
        }

        $sheetFormulaCount = 0
        $sheetCellCount = 0

        for ($rowIndex = $rangeInfo.first_row; $rowIndex -le $rangeInfo.last_row; $rowIndex++) {
            for ($columnIndex = $rangeInfo.first_column; $columnIndex -le $rangeInfo.last_column; $columnIndex++) {
                $cell = $null
                $mergeArea = $null
                try {
                    $cell = $sheet.Cells.Item($rowIndex, $columnIndex)
                    $address = [string]$cell.Address($false, $false)
                    $hasFormula = [bool]$cell.HasFormula
                    $formula = if ($hasFormula) { [string]$cell.Formula } else { $null }
                    $formula2 = if ($hasFormula) { Get-CellFormula2 -Cell $cell } else { $null }
                    $mergeAreaAddress = $null
                    $isMergeAnchor = $false

                    if ([bool]$cell.MergeCells) {
                        $mergeArea = $cell.MergeArea
                        $mergeAreaAddress = [string]$mergeArea.Address($false, $false)
                        $isMergeAnchor = ([int]$mergeArea.Row -eq $rowIndex -and [int]$mergeArea.Column -eq $columnIndex)

                        if ($sheetMergedKeys.Add($mergeAreaAddress)) {
                            $anchorAddress = Convert-CoordinateToA1 -Row ([int]$mergeArea.Row) -Column ([int]$mergeArea.Column)
                            $mergeRecord = [ordered]@{
                                sheet = [string]$sheet.Name
                                range = $mergeAreaAddress
                                anchor = $anchorAddress
                            }
                            $sheetMergedRanges.Add($mergeRecord)
                            if ($globalMergedKeys.Add(("{0}|{1}" -f [string]$sheet.Name, $mergeAreaAddress))) {
                                $mergedRanges.Add($mergeRecord)
                            }
                        }
                    }

                    $cellRecord = [ordered]@{
                        sheet = [string]$sheet.Name
                        address = $address
                        row = $rowIndex
                        column = $columnIndex
                        value2 = Convert-VariantValue -Value $cell.Value2
                        text = [string]$cell.Text
                        formula = $formula
                        formula2 = $formula2
                        has_formula = $hasFormula
                        number_format = [string]$cell.NumberFormat
                        merge_area = $mergeAreaAddress
                        is_merge_anchor = $isMergeAnchor
                        comment = Get-CellCommentText -Cell $cell
                        comment_threaded = Get-CellThreadedComment -Cell $cell
                        hyperlink = Get-CellHyperlink -Cell $cell
                    }

                    $cells.Add($cellRecord)
                    $sheetCellCount++
                    $totalCellCount++
                    if ($hasFormula) {
                        $sheetFormulaCount++
                        $totalFormulaCount++
                    }

                    if ($CollectStyles) {
                        try {
                            $styleRecord = Get-StyleRecord -Cell $cell
                            $styles.Add([ordered]@{
                                sheet = [string]$sheet.Name
                                address = $address
                                fill_color = $styleRecord.fill_color
                                font_color = $styleRecord.font_color
                                horizontal_alignment = $styleRecord.horizontal_alignment
                                vertical_alignment = $styleRecord.vertical_alignment
                                wrap_text = $styleRecord.wrap_text
                                borders = $styleRecord.borders
                            })
                        }
                        catch {
                            Add-WarningMessage -Warnings $warnings -Message ("styles.json export skipped for {0}!{1}: {2}" -f [string]$sheet.Name, $address, $_.Exception.Message)
                        }
                    }
                }
                finally {
                    if ($null -ne $mergeArea) {
                        Release-ComReference $mergeArea
                    }
                    if ($null -ne $cell) {
                        Release-ComReference $cell
                    }
                }
            }
        }

        $sheets.Add([ordered]@{
            sheet_name = [string]$sheet.Name
            sheet_index = [int]$sheet.Index
            visible = [int]$sheet.Visible
            used_range = $rangeInfo
            freeze_panes = $freezePanes
            hidden_rows = $hiddenRows
            hidden_columns = $hiddenColumns
            row_heights = $rowHeights
            column_widths = $columnWidths
            cell_count = $sheetCellCount
            formula_count = $sheetFormulaCount
            merged_ranges = $sheetMergedRanges
        })

        if ($null -ne $usedRange) {
            Release-ComReference $usedRange
            $usedRange = $null
        }

        Release-ComReference $sheet
    }

    $workbookPayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        workbook = [ordered]@{
            name = [string]$workbook.Name
            path = $resolvedExcelPath
            extension = [System.IO.Path]::GetExtension($resolvedExcelPath)
            sheet_count = [int]$workbook.Worksheets.Count
            has_vba = @('.xlsm', '.xlam') -contains ([System.IO.Path]::GetExtension($resolvedExcelPath).ToLowerInvariant())
        }
        sheets = $sheets
        cells = $cells
        merged_ranges = $mergedRanges
    }

    $stylePayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        styles = $styles
    }

    $styleStatus = if (-not $CollectStyles) { 'skipped' } elseif ($styles.Count -gt 0) { 'generated' } else { 'empty' }
    $status = if ($warnings.Count -gt 0) { 'warning' } else { 'success' }

    $manifestPayload = [ordered]@{
        generated_at = Get-TimestampJst
        generator = 'Excel2LLM PowerShell Extractor'
        status = $status
        warnings = $warnings
        workbook_path = $resolvedExcelPath
        output_directory = [System.IO.Path]::GetFullPath($OutputDir)
        sheet_count = [int]$workbook.Worksheets.Count
        cell_count = $totalCellCount
        formula_count = $totalFormulaCount
        merged_range_count = $mergedRanges.Count
        style_export_status = $styleStatus
        verify_status = 'not_run'
    }

    Write-JsonFile -Data $workbookPayload -Path $workbookJsonPath
    Write-JsonFile -Data $stylePayload -Path $stylesJsonPath
    Write-JsonFile -Data $manifestPayload -Path $manifestJsonPath

    Write-Host "Extracted workbook.json -> $workbookJsonPath"
    Write-Host "Extracted styles.json   -> $stylesJsonPath"
    Write-Host "Extracted manifest.json -> $manifestJsonPath"
}
catch {
    throw "extract_excel.ps1 line $($_.InvocationInfo.ScriptLineNumber): $($_.Exception.Message)"
}
finally {
    if ($null -ne $usedRange) {
        Release-ComReference $usedRange
    }
    if ($null -ne $workbook) {
        try {
            $workbook.Close($false)
        }
        catch {
        }
        Release-ComReference $workbook
    }
    if ($null -ne $excel) {
        try {
            $excel.Quit()
        }
        catch {
        }
        Release-ComReference $excel
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
