[CmdletBinding()]
param(
    [string]$OutputRoot,
    [ValidateSet('all', 'mechanical', 'accounting')]
    [string]$Scenario = 'all',
    [ValidateSet('original', 'improved', 'both')]
    [string]$Variant = 'both',
    [int]$MaxCells = 180
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $OutputRoot) {
    $OutputRoot = Join-Path (Join-Path (Get-ProjectRoot) 'output') ('domain_acceptance_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
}

Ensure-Directory -Path $OutputRoot
$samplesDir = Join-Path $OutputRoot 'samples'
Ensure-Directory -Path $samplesDir

& (Join-Path $PSScriptRoot 'create_domain_sample_workbooks.ps1') -OutputDir $samplesDir -Scenario $Scenario -Variant $Variant

$results = [System.Collections.Generic.List[object]]::new()
$files = Get-ChildItem -LiteralPath $samplesDir -Filter '*.xlsx' | Sort-Object Name

foreach ($file in $files) {
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($file.Name)
    $parts = $baseName.Split('_')
    $scenarioName = $parts[0]
    $variantName = $parts[1]
    $scenarioOutputDir = Join-Path $OutputRoot $baseName
    Ensure-Directory -Path $scenarioOutputDir

    $extractDir = Join-Path $scenarioOutputDir 'extract'
    $promptDir = Join-Path $scenarioOutputDir 'prompts'
    Ensure-Directory -Path $extractDir
    Ensure-Directory -Path $promptDir

    $startedAt = Get-Date
    & (Join-Path $PSScriptRoot 'extract_excel.ps1') -ExcelPath $file.FullName -OutputDir $extractDir -CollectStyles
    & (Join-Path $PSScriptRoot 'pack_for_llm.ps1') -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -OutputPath (Join-Path $extractDir 'llm_package.jsonl') -ChunkBy range -MaxCells $MaxCells -IncludeStyles -StylesJsonPath (Join-Path $extractDir 'styles.json')
    & (Join-Path $PSScriptRoot 'excel_verify.ps1') -ExcelPath $file.FullName -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -OutputDir $extractDir
    & (Join-Path $PSScriptRoot 'export_prompt_bundle.ps1') -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario $scenarioName -OutputDir $promptDir
    $finishedAt = Get-Date

    $manifest = Get-Content -LiteralPath (Join-Path $extractDir 'manifest.json') -Raw | ConvertFrom-Json
    $verify = Get-Content -LiteralPath (Join-Path $extractDir 'verify_report.json') -Raw | ConvertFrom-Json
    $promptManifest = Get-Content -LiteralPath (Join-Path $promptDir 'prompt_bundle_manifest.json') -Raw | ConvertFrom-Json
    $chunkCount = @(Get-Content -LiteralPath (Join-Path $extractDir 'llm_package.jsonl') | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count

    [void]$results.Add([ordered]@{
        scenario = $scenarioName
        variant = $variantName
        workbook_path = $file.FullName
        extract_dir = $extractDir
        prompt_dir = $promptDir
        started_at = $startedAt.ToString("yyyy-MM-dd HH:mm:ss")
        finished_at = $finishedAt.ToString("yyyy-MM-dd HH:mm:ss")
        duration_seconds = [Math]::Round(($finishedAt - $startedAt).TotalSeconds, 2)
        sheet_count = [int]$manifest.sheet_count
        cell_count = [int]$manifest.cell_count
        formula_count = [int]$manifest.formula_count
        merged_range_count = [int]$manifest.merged_range_count
        warning_count = @($manifest.warnings).Count
        mismatch_count = [int]$verify.mismatch_count
        chunk_count = $chunkCount
        prompt_count = [int]$promptManifest.prompt_count
    })
}

$comparisons = [System.Collections.Generic.List[object]]::new()
foreach ($scenarioName in (@($results.ToArray() | ForEach-Object { $_.scenario } | Select-Object -Unique))) {
    $original = $results | Where-Object { $_.scenario -eq $scenarioName -and $_.variant -eq 'original' } | Select-Object -First 1
    $improved = $results | Where-Object { $_.scenario -eq $scenarioName -and $_.variant -eq 'improved' } | Select-Object -First 1
    if ($null -eq $original -or $null -eq $improved) {
        continue
    }

    [void]$comparisons.Add([ordered]@{
        scenario = $scenarioName
        original_formula_count = $original.formula_count
        improved_formula_count = $improved.formula_count
        original_chunk_count = $original.chunk_count
        improved_chunk_count = $improved.chunk_count
        original_warning_count = $original.warning_count
        improved_warning_count = $improved.warning_count
        original_mismatch_count = $original.mismatch_count
        improved_mismatch_count = $improved.mismatch_count
    })
}

$summary = [ordered]@{
    generated_at = Get-TimestampJst
    output_root = $OutputRoot
    scenario = $Scenario
    variant = $Variant
    result_count = $results.Count
    results = @($results)
    comparisons = @($comparisons)
}

Write-JsonFile -Data $summary -Path (Join-Path $OutputRoot 'scenario_summary.json')
Write-Host "Domain acceptance completed -> $OutputRoot"
