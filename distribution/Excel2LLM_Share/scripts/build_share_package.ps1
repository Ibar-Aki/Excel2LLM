[CmdletBinding()]
param(
    [string]$OutputDir
)

. (Join-Path $PSScriptRoot 'common.ps1')

$projectRoot = Get-ProjectRoot

if (-not $OutputDir) {
    $OutputDir = Join-Path $projectRoot 'distribution\Excel2LLM_Share'
}

$resolvedOutputDir = [System.IO.Path]::GetFullPath($OutputDir)

if (Test-Path -LiteralPath $resolvedOutputDir) {
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
    'README.md',
    'run_build_share_package.bat',
    'run_domain_acceptance.bat',
    'run_extract.bat',
    'run_pack.bat',
    'run_rebuild.bat',
    'run_self_test.bat',
    'run_verify.bat',
    'docs\FORMAT.md',
    'docs\LLM_PROMPT_FORMATS.md',
    'docs\MANUAL.md',
    'docs\SHARE_PACKAGE.md',
    'docs\USE_CASES.md',
    'docs\USER_GUIDE.md',
    'docs\VBA_HELPER.md',
    'output\.gitkeep',
    'samples\.gitkeep',
    'scripts\build_share_package.ps1',
    'scripts\common.ps1',
    'scripts\create_domain_sample_workbooks.ps1',
    'scripts\create_sample_workbook.ps1',
    'scripts\excel_verify.ps1',
    'scripts\export_prompt_bundle.ps1',
    'scripts\extract_excel.ps1',
    'scripts\pack_for_llm.ps1',
    'scripts\rebuild_excel.ps1',
    'scripts\run_domain_acceptance.ps1',
    'scripts\run_self_test.ps1',
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
    source_project_root = $projectRoot
    output_directory = $resolvedOutputDir
    file_count = $copiedFiles.Count
    files = @($copiedFiles)
}

Write-JsonFile -Data $manifest -Path (Join-Path $resolvedOutputDir 'share_manifest.json')
Write-Host "Built share package -> $resolvedOutputDir"
