[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string]$WorkbookJsonPath,
    [string]$OutputPath,
    [ValidateSet('sheet', 'range')]
    [string]$ChunkBy = 'sheet',
    [int]$MaxCells = 500,
    [switch]$IncludeStyles,
    [string]$StylesJsonPath
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $OutputPath) {
    $OutputPath = Join-Path (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'output') 'llm_package.jsonl'
}

if (-not $StylesJsonPath) {
    $StylesJsonPath = Join-Path (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'output') 'styles.json'
}

function Get-ChunkRange {
    param(
        [Parameter(Mandatory)]
        [object[]]$ChunkCells
    )

    $minRow = ($ChunkCells | Measure-Object -Property row -Minimum).Minimum
    $maxRow = ($ChunkCells | Measure-Object -Property row -Maximum).Maximum
    $minColumn = ($ChunkCells | Measure-Object -Property column -Minimum).Minimum
    $maxColumn = ($ChunkCells | Measure-Object -Property column -Maximum).Maximum

    $start = Convert-CoordinateToA1 -Row $minRow -Column $minColumn
    $end = Convert-CoordinateToA1 -Row $maxRow -Column $maxColumn
    return '{0}:{1}' -f $start, $end
}

function Get-ChunkPayload {
    param(
        [Parameter(Mandatory)]
        [string]$SheetName,
        [Parameter(Mandatory)]
        [string]$ChunkRange,
        [Parameter(Mandatory)]
        [object[]]$ChunkCells,
        [Parameter(Mandatory)]
        [hashtable]$StyleLookup,
        [switch]$IncludeStyles
    )

    $payloadCells = foreach ($cell in $ChunkCells) {
        $entry = [ordered]@{
            address = $cell.address
            row = [int]$cell.row
            column = [int]$cell.column
            value2 = $cell.value2
            text = $cell.text
            formula = $cell.formula
            formula2 = $cell.formula2
            has_formula = [bool]$cell.has_formula
            number_format = $cell.number_format
            merge_area = $cell.merge_area
            is_merge_anchor = [bool]$cell.is_merge_anchor
            comment = $cell.comment
            comment_threaded = $cell.comment_threaded
            hyperlink = $cell.hyperlink
        }

        if ($IncludeStyles) {
            $styleKey = '{0}|{1}' -f $SheetName, $cell.address
            if ($StyleLookup.ContainsKey($styleKey)) {
                $entry['style'] = $StyleLookup[$styleKey]
            }
        }

        $entry
    }

    return [ordered]@{
        sheet_name = $SheetName
        range = $ChunkRange
        cell_count = $ChunkCells.Count
        cells = $payloadCells
    }
}

$resolvedWorkbookJsonPath = Resolve-AbsolutePath -Path $WorkbookJsonPath
$workbookData = Get-Content -LiteralPath $resolvedWorkbookJsonPath -Raw | ConvertFrom-Json
$styleLookup = @{}

if ($IncludeStyles -and (Test-Path -LiteralPath $StylesJsonPath)) {
    $stylesData = Get-Content -LiteralPath $StylesJsonPath -Raw | ConvertFrom-Json
    foreach ($style in $stylesData.styles) {
        $styleLookup['{0}|{1}' -f $style.sheet, $style.address] = [ordered]@{
            fill_color = $style.fill_color
            font_color = $style.font_color
            horizontal_alignment = $style.horizontal_alignment
            vertical_alignment = $style.vertical_alignment
            wrap_text = $style.wrap_text
            borders = $style.borders
        }
    }
}

$chunks = New-Object System.Collections.Generic.List[object]
$chunkIndex = 0

foreach ($sheet in $workbookData.sheets) {
    $sheetCells = @($workbookData.cells | Where-Object { $_.sheet -eq $sheet.sheet_name } | Sort-Object row, column)
    if ($sheetCells.Count -eq 0) {
        continue
    }

    for ($offset = 0; $offset -lt $sheetCells.Count; $offset += $MaxCells) {
        $upperBound = [Math]::Min($offset + $MaxCells - 1, $sheetCells.Count - 1)
        $chunkCells = @($sheetCells[$offset..$upperBound])
        $chunkRange = Get-ChunkRange -ChunkCells $chunkCells
        $payload = Get-ChunkPayload -SheetName $sheet.sheet_name -ChunkRange $chunkRange -ChunkCells $chunkCells -StyleLookup $styleLookup -IncludeStyles:$IncludeStyles
        $payloadJson = $payload | ConvertTo-Json -Depth 40 -Compress
        $formulaCells = @($chunkCells | Where-Object { $_.has_formula } | ForEach-Object { $_.address })

        $chunks.Add([ordered]@{
            chunk_id = ('{0}-{1:D4}' -f $sheet.sheet_name, $chunkIndex)
            sheet_name = $sheet.sheet_name
            range = $chunkRange
            cell_addresses = @($chunkCells | ForEach-Object { $_.address })
            payload = $payload
            formula_cells = $formulaCells
            token_estimate = [Math]::Ceiling($payloadJson.Length / 4)
            includes_styles = [bool]$IncludeStyles
        })

        $chunkIndex++
    }
}

Ensure-Directory -Path (Split-Path -Path $OutputPath -Parent)
Write-JsonLineFile -Items $chunks -Path $OutputPath
Write-Host "Packed llm_package.jsonl -> $OutputPath"
