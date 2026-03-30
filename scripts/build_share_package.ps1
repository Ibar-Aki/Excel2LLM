[CmdletBinding()]
param(
    [string]$OutputDir,
    [switch]$AllowOutsideDistribution,
    [switch]$ForceCleanOutputDir
)

. (Join-Path $PSScriptRoot 'common.ps1')

$projectRoot = Get-ProjectRoot

if (-not $OutputDir) {
    $OutputDir = Join-Path $projectRoot 'distribution\Excel2LLM_Share'
}

$distributionRoot = Join-Path $projectRoot 'distribution'
$resolvedOutputDir = Get-NormalizedFullPath -Path $OutputDir
$resolvedDistributionRoot = Get-NormalizedFullPath -Path $distributionRoot
$isWithinDistribution = Test-PathWithinDirectory -Path $resolvedOutputDir -DirectoryPath $resolvedDistributionRoot

if (-not $isWithinDistribution -and -not $AllowOutsideDistribution) {
    throw "OutputDir outside distribution is blocked by default: $resolvedOutputDir. Use -AllowOutsideDistribution to continue."
}

if (Test-Path -LiteralPath $resolvedOutputDir) {
    if (-not $isWithinDistribution -and -not $ForceCleanOutputDir) {
        throw "Cleaning an existing OutputDir outside distribution requires -ForceCleanOutputDir: $resolvedOutputDir"
    }

    Remove-Item -LiteralPath $resolvedOutputDir -Recurse -Force
}

Ensure-Directory -Path $resolvedOutputDir

$directories = @(
    'docs',
    'output',
    'samples',
    'scripts',
    'templates'
)

foreach ($directory in $directories) {
    Ensure-Directory -Path (Join-Path $resolvedOutputDir $directory)
}

$filesToCopy = @(
    'Excel2LLM.bat',
    'GETTING_STARTED.md',
    'README.md',
    'docs\README.md',
    'docs\guides\USE_CASES.md',
    'docs\reference\EXTRACTION_COVERAGE.md',
    'docs\reference\FORMAT.md',
    'docs\reference\LLM_PROMPT_FORMATS.md',
    'docs\reference\VBA_HELPER.md',
    'output\.gitkeep',
    'samples\.gitkeep',
    'scripts\common.ps1',
    'scripts\create_sample_workbook.ps1',
    'scripts\excel_verify.ps1',
    'scripts\export_prompt_bundle.ps1',
    'scripts\extract_excel.ps1',
    'scripts\invoke_excel2llm.ps1',
    'scripts\launch_menu.ps1',
    'scripts\macro_extract.ps1',
    'scripts\pack_for_llm.ps1',
    'scripts\preflight_excel.ps1',
    'scripts\rebuild_excel.ps1',
    'scripts\run_all.ps1',
    'scripts\run_prompt_bundle.ps1',
    'scripts\run_self_test.ps1',
    'scripts\show_usage.ps1',
    'tools\user\run_all.bat',
    'tools\user\run_prompt_bundle.bat',
    'tools\user\run_self_test.bat',
    'tools\advanced\run_extract.bat',
    'tools\advanced\run_macro_extract.bat',
    'tools\advanced\run_pack.bat',
    'tools\advanced\run_preflight.bat',
    'tools\advanced\run_rebuild.bat',
    'tools\advanced\run_verify.bat',
    'templates\Excel2LLM_Helper.bas'
)

$copiedFiles = [System.Collections.Generic.List[string]]::new()

foreach ($relativePath in $filesToCopy) {
    $sourcePath = Join-Path $projectRoot $relativePath
    if (-not (Test-Path -LiteralPath $sourcePath)) {
        throw "Share package source file was not found: $sourcePath"
    }

    $destinationPath = Join-Path $resolvedOutputDir $relativePath
    Ensure-Directory -Path (Split-Path -Path $destinationPath -Parent)
    Copy-Item -LiteralPath $sourcePath -Destination $destinationPath -Force
    [void]$copiedFiles.Add($relativePath)
}

$manifest = [ordered]@{
    generated_at = Get-TimestampJst
    generator = 'Excel2LLM Share Package Builder'
    package_name = [string](Split-Path -Path $resolvedOutputDir -Leaf)
    file_count = $copiedFiles.Count
    files = @($copiedFiles)
}

Write-JsonFile -Data $manifest -Path (Join-Path $resolvedOutputDir 'share_manifest.json')
Write-Host "Built share package -> $resolvedOutputDir"
