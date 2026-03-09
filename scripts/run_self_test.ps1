[CmdletBinding()]
param()

. (Join-Path $PSScriptRoot 'common.ps1')

$projectRoot = Split-Path -Path $PSScriptRoot -Parent
$samplesDir = Join-Path $projectRoot 'samples'
$outputDir = Join-Path $projectRoot 'output'

& (Join-Path $PSScriptRoot 'create_sample_workbook.ps1') -OutputDir $samplesDir
& (Join-Path $PSScriptRoot 'extract_excel.ps1') -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $outputDir
& (Join-Path $PSScriptRoot 'pack_for_llm.ps1') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath (Join-Path $outputDir 'llm_package.jsonl') -ChunkBy range -MaxCells 250
& (Join-Path $PSScriptRoot 'excel_verify.ps1') -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputDir $outputDir

$workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
$manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json
$verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json
$jsonlLines = @(Get-Content -LiteralPath (Join-Path $outputDir 'llm_package.jsonl'))

if ($workbookJson.sheets.Count -lt 3) {
    throw 'Expected at least three worksheets in the sample workbook.'
}

if (($workbookJson.cells | Where-Object { $_.has_formula }).Count -lt 3) {
    throw 'Expected formula cells were not extracted.'
}

$firstFormulaCell = $workbookJson.cells | Where-Object { $_.has_formula } | Select-Object -First 1
if ($null -eq $firstFormulaCell.PSObject.Properties['formula2']) {
    throw 'formula2 field is missing from extracted cells.'
}

if (($workbookJson.merged_ranges | Where-Object { $_.sheet -eq 'Summary' }).Count -lt 1) {
    throw 'Merged range extraction failed.'
}

if ($jsonlLines.Count -lt 2) {
    throw 'Expected multiple JSONL chunks.'
}

if (-not @('success', 'warning') -contains [string]$manifest.status) {
    throw 'manifest.json status is invalid.'
}

if (-not @('success', 'warning') -contains [string]$verifyReport.status) {
    throw 'verify_report.json status is invalid.'
}

Write-Host 'Self-test completed successfully.'
