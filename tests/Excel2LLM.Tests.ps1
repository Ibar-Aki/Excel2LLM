$projectRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$extractScript = Join-Path $projectRoot 'scripts\extract_excel.ps1'
$packScript = Join-Path $projectRoot 'scripts\pack_for_llm.ps1'
$verifyScript = Join-Path $projectRoot 'scripts\excel_verify.ps1'
$rebuildScript = Join-Path $projectRoot 'scripts\rebuild_excel.ps1'
$sampleScript = Join-Path $projectRoot 'scripts\create_sample_workbook.ps1'
$domainSampleScript = Join-Path $projectRoot 'scripts\create_domain_sample_workbooks.ps1'
$promptBundleScript = Join-Path $projectRoot 'scripts\export_prompt_bundle.ps1'
$acceptanceScript = Join-Path $projectRoot 'scripts\run_domain_acceptance.ps1'

Describe 'Excel2LLM integration tests' {
    It 'extracts workbook metadata, formulas, merge information, and xlsm VBA metadata' {
        $workspace = New-TestWorkspace -Name 'extract'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'

        & $sampleScript -OutputDir $samplesDir
        & $extractScript -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $outputDir

        $workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json

        $workbookJson.workbook.sheet_count | Should Be 3
        $manifest.formula_count | Should BeGreaterThan 3
        ($workbookJson.merged_ranges | Where-Object { $_.sheet -eq 'Summary' }).Count | Should Be 1
        (@(($workbookJson.sheets | Where-Object { $_.sheet_name -eq 'Summary' }).hidden_rows) -contains 7) | Should Be $true
        (@(($workbookJson.sheets | Where-Object { $_.sheet_name -eq 'Summary' }).hidden_columns) -contains 'E') | Should Be $true

        $formulaCell = $workbookJson.cells | Where-Object { $_.sheet -eq 'Calc' -and $_.address -eq 'A3' } | Select-Object -First 1
        $formulaCell.formula | Should Be '=SUM(A1:A2)'
        $formulaCell.formula2 | Should Be '=SUM(A1:A2)'

        & $extractScript -ExcelPath (Join-Path $samplesDir 'sample.xlsm') -OutputDir $outputDir
        $xlsmWorkbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $xlsmWorkbookJson.workbook.has_vba | Should Be $true
    }

    It 'differentiates sheet chunking from range chunking' {
        $workspace = New-TestWorkspace -Name 'chunking'
        $bookPath = Join-Path $workspace 'grid.xlsx'
        $outputDir = Join-Path $workspace 'output'
        $sheetJsonl = Join-Path $workspace 'sheet.jsonl'
        $rangeJsonl = Join-Path $workspace 'range.jsonl'

        New-MiniWorkbook -Path $bookPath
        & $extractScript -ExcelPath $bookPath -OutputDir $outputDir
        & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath $sheetJsonl -ChunkBy sheet -MaxCells 5
        & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath $rangeJsonl -ChunkBy range -MaxCells 5

        $sheetChunks = @(Get-Content -LiteralPath $sheetJsonl | ForEach-Object { $_ | ConvertFrom-Json })
        $rangeChunks = @(Get-Content -LiteralPath $rangeJsonl | ForEach-Object { $_ | ConvertFrom-Json })

        $sheetChunks[0].range | Should Be 'A1:D1'
        $sheetChunks[0].cell_addresses.Count | Should Be 4
        $rangeChunks[0].range | Should Be 'A1:D2'
        $rangeChunks[0].cell_addresses.Count | Should Be 5
    }

    It 'includes style payload only when styles are explicitly collected and requested' {
        $workspace = New-TestWorkspace -Name 'styles'
        $bookPath = Join-Path $workspace 'styles.xlsx'
        $outputDir = Join-Path $workspace 'output'
        $jsonlPath = Join-Path $workspace 'styles.jsonl'

        New-MiniWorkbook -Path $bookPath -IncludeStyles
        & $extractScript -ExcelPath $bookPath -OutputDir $outputDir -CollectStyles
        & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath $jsonlPath -ChunkBy sheet -MaxCells 20 -IncludeStyles -StylesJsonPath (Join-Path $outputDir 'styles.json')

        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json
        $styles = Get-Content -LiteralPath (Join-Path $outputDir 'styles.json') -Raw | ConvertFrom-Json
        $chunks = @(Get-Content -LiteralPath $jsonlPath | ForEach-Object { $_ | ConvertFrom-Json })
        $styledCell = $chunks[0].payload.cells | Where-Object { $_.address -eq 'A1' } | Select-Object -First 1

        $manifest.style_export_status | Should Be 'generated'
        $styles.styles.Count | Should BeGreaterThan 0
        $styledCell.style.fill_color | Should Be '#FF0000'
        $styledCell.style.wrap_text | Should Be $true
    }

    It 'reports mismatches when workbook.json is tampered after extraction' {
        $workspace = New-TestWorkspace -Name 'verify'
        $bookPath = Join-Path $workspace 'verify.xlsx'
        $outputDir = Join-Path $workspace 'output'

        New-MiniWorkbook -Path $bookPath
        & $extractScript -ExcelPath $bookPath -OutputDir $outputDir

        $workbookPath = Join-Path $outputDir 'workbook.json'
        $tampered = Get-Content -LiteralPath $workbookPath -Raw | ConvertFrom-Json
        $targetCell = $tampered.cells | Where-Object { $_.address -eq 'A1' } | Select-Object -First 1
        $targetCell.text = 'BROKEN'
        Write-JsonFile -Data $tampered -Path $workbookPath

        & $verifyScript -ExcelPath $bookPath -WorkbookJsonPath $workbookPath -OutputDir $outputDir

        $verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json
        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json

        $verifyReport.status | Should Be 'warning'
        $verifyReport.mismatch_count | Should BeGreaterThan 0
        $manifest.verify_status | Should Be 'warning'
    }

    It 'round-trips workbook.json back into xlsx while preserving formulas, merges, comments, links, hidden state, and freeze panes' {
        $workspace = New-TestWorkspace -Name 'rebuild-roundtrip'
        $samplesDir = Join-Path $workspace 'samples'
        $sourceOutputDir = Join-Path $workspace 'source-output'
        $rebuiltDir = Join-Path $workspace 'rebuilt'
        $roundTripOutputDir = Join-Path $workspace 'roundtrip-output'
        $rebuiltPath = Join-Path $rebuiltDir 'sample_roundtrip.xlsx'

        & $sampleScript -OutputDir $samplesDir
        & $extractScript -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $sourceOutputDir
        & $rebuildScript -WorkbookJsonPath (Join-Path $sourceOutputDir 'workbook.json') -OutputPath $rebuiltPath -Overwrite
        & $extractScript -ExcelPath $rebuiltPath -OutputDir $roundTripOutputDir

        $source = Get-Content -LiteralPath (Join-Path $sourceOutputDir 'workbook.json') -Raw | ConvertFrom-Json
        $roundTrip = Get-Content -LiteralPath (Join-Path $roundTripOutputDir 'workbook.json') -Raw | ConvertFrom-Json
        $rebuildReport = Get-Content -LiteralPath (Join-Path $rebuiltDir 'rebuild_report.json') -Raw | ConvertFrom-Json

        $source.workbook.sheet_count | Should Be $roundTrip.workbook.sheet_count
        $source.merged_ranges.Count | Should Be $roundTrip.merged_ranges.Count

        $sourceSummary = $source.sheets | Where-Object { $_.sheet_name -eq 'Summary' } | Select-Object -First 1
        $roundTripSummary = $roundTrip.sheets | Where-Object { $_.sheet_name -eq 'Summary' } | Select-Object -First 1
        $roundTripSummary.freeze_panes.enabled | Should Be $true
        $roundTripSummary.freeze_panes.split_row | Should Be 1
        $roundTripSummary.freeze_panes.split_column | Should Be 1
        (@($roundTripSummary.hidden_rows) -contains 7) | Should Be $true
        (@($roundTripSummary.hidden_columns) -contains 'E') | Should Be $true
        $roundTripSummary.formula_count | Should Be $sourceSummary.formula_count

        $formulaCell = $roundTrip.cells | Where-Object { $_.sheet -eq 'Calc' -and $_.address -eq 'A3' } | Select-Object -First 1
        $formulaCell.formula | Should Be '=SUM(A1:A2)'

        $hyperlinkCell = $roundTrip.cells | Where-Object { $_.sheet -eq 'Summary' -and $_.address -eq 'A8' } | Select-Object -First 1
        $hyperlinkCell.hyperlink.address | Should Match '^https://example\.com/?$'

        $commentCell = $roundTrip.cells | Where-Object { $_.sheet -eq 'Summary' -and $_.address -eq 'B8' } | Select-Object -First 1
        $commentCell.comment | Should Be 'Legacy comment'

        $rebuildReport.status | Should Be 'success'
        $rebuildReport.restored_sheets | Should Be 3
    }

    It 'rebuilds xlsm-derived workbook.json as xlsx and reports VBA was not restored' {
        $workspace = New-TestWorkspace -Name 'rebuild-xlsm'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'
        $rebuiltPath = Join-Path $workspace 'rebuilt\sample_from_xlsm.xlsx'

        & $sampleScript -OutputDir $samplesDir
        & $extractScript -ExcelPath (Join-Path $samplesDir 'sample.xlsm') -OutputDir $outputDir
        & $rebuildScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath $rebuiltPath -Overwrite

        $rebuildReport = Get-Content -LiteralPath (Join-Path (Split-Path -Path $rebuiltPath -Parent) 'rebuild_report.json') -Raw | ConvertFrom-Json

        (Test-Path -LiteralPath $rebuiltPath) | Should Be $true
        [System.IO.Path]::GetExtension($rebuildReport.output_path) | Should Be '.xlsx'
        $rebuildReport.source_has_vba | Should Be $true
        ($rebuildReport.warnings -join ' ') | Should Match 'VBA'
    }

    It 'restores styles only when styles.json is available to the rebuild step' {
        $workspace = New-TestWorkspace -Name 'rebuild-styles'
        $bookPath = Join-Path $workspace 'styles.xlsx'
        $extractDir = Join-Path $workspace 'extract'
        $withStylesDir = Join-Path $workspace 'with-styles'
        $withoutStylesDir = Join-Path $workspace 'without-styles'
        $withStylesRoundTripDir = Join-Path $workspace 'with-styles-output'
        $withoutStylesRoundTripDir = Join-Path $workspace 'without-styles-output'

        New-MiniWorkbook -Path $bookPath -IncludeStyles
        & $extractScript -ExcelPath $bookPath -OutputDir $extractDir -CollectStyles

        $isolatedWorkbookJson = Join-Path $withoutStylesDir 'workbook.json'
        Ensure-Directory -Path $withoutStylesDir
        Copy-Item -LiteralPath (Join-Path $extractDir 'workbook.json') -Destination $isolatedWorkbookJson

        & $rebuildScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -StylesJsonPath (Join-Path $extractDir 'styles.json') -OutputPath (Join-Path $withStylesDir 'rebuilt.xlsx') -Overwrite
        & $extractScript -ExcelPath (Join-Path $withStylesDir 'rebuilt.xlsx') -OutputDir $withStylesRoundTripDir -CollectStyles

        & $rebuildScript -WorkbookJsonPath $isolatedWorkbookJson -OutputPath (Join-Path $withoutStylesDir 'rebuilt.xlsx') -Overwrite
        & $extractScript -ExcelPath (Join-Path $withoutStylesDir 'rebuilt.xlsx') -OutputDir $withoutStylesRoundTripDir -CollectStyles

        $withStyles = Get-Content -LiteralPath (Join-Path $withStylesRoundTripDir 'styles.json') -Raw | ConvertFrom-Json
        $withoutStyles = Get-Content -LiteralPath (Join-Path $withoutStylesRoundTripDir 'styles.json') -Raw | ConvertFrom-Json
        $withStyledCell = $withStyles.styles | Where-Object { $_.sheet -eq 'Grid' -and $_.address -eq 'A1' } | Select-Object -First 1
        $withoutStyledCell = $withoutStyles.styles | Where-Object { $_.sheet -eq 'Grid' -and $_.address -eq 'A1' } | Select-Object -First 1

        $withStyledCell.fill_color | Should Be '#FF0000'
        $withStyledCell.wrap_text | Should Be $true
        $withoutStyledCell.fill_color | Should Not Be '#FF0000'
    }

    It 'fails fast when workbook.json is missing' {
        $workspace = New-TestWorkspace -Name 'rebuild-invalid'
        $rebuiltPath = Join-Path $workspace 'rebuilt\invalid.xlsx'
        $missingPath = Join-Path $workspace 'missing-workbook.json'

        $didThrow = $false
        try {
            & $rebuildScript -WorkbookJsonPath $missingPath -OutputPath $rebuiltPath -Overwrite
        }
        catch {
            $didThrow = $true
        }

        $didThrow | Should Be $true
    }

    It 'creates domain samples and exports scenario-specific prompt bundles' {
        $workspace = New-TestWorkspace -Name 'domain-prompts'
        $samplesDir = Join-Path $workspace 'samples'
        $extractDir = Join-Path $workspace 'extract'
        $promptDir = Join-Path $workspace 'prompts'
        $keepFile = Join-Path $promptDir 'keep.txt'

        & $domainSampleScript -OutputDir $samplesDir -Scenario mechanical -Variant original
        & $extractScript -ExcelPath (Join-Path $samplesDir 'mechanical_original.xlsx') -OutputDir $extractDir -CollectStyles
        & $packScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -OutputPath (Join-Path $extractDir 'llm_package.jsonl') -ChunkBy range -MaxCells 120 -IncludeStyles -StylesJsonPath (Join-Path $extractDir 'styles.json')
        & $promptBundleScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario mechanical -OutputDir $promptDir
        Set-Content -LiteralPath $keepFile -Value 'do not delete'
        & $promptBundleScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario mechanical -OutputDir $promptDir

        $manifest = Get-Content -LiteralPath (Join-Path $promptDir 'prompt_bundle_manifest.json') -Raw | ConvertFrom-Json
        $promptFiles = @(Get-ChildItem -LiteralPath $promptDir -Filter '*_mechanical_*.txt')

        (Test-Path -LiteralPath (Join-Path $samplesDir 'mechanical_original.xlsx')) | Should Be $true
        $manifest.scenario | Should Be 'mechanical'
        $manifest.prompt_count | Should BeGreaterThan 0
        $manifest.prompts[0].sheet_name | Should Be 'Calc'
        $promptFiles.Count | Should Be $manifest.prompt_count
        (Test-Path -LiteralPath $keepFile) | Should Be $true
        (Test-Path -LiteralPath $manifest.prompts[0].path) | Should Be $true
    }

    It 'runs domain acceptance workflow for accounting original sample' {
        $workspace = New-TestWorkspace -Name 'domain-acceptance'

        & $acceptanceScript -OutputRoot $workspace -Scenario accounting -Variant original -MaxCells 120

        $summary = Get-Content -LiteralPath (Join-Path $workspace 'scenario_summary.json') -Raw | ConvertFrom-Json
        $result = $summary.results | Select-Object -First 1

        $summary.result_count | Should Be 1
        $result.scenario | Should Be 'accounting'
        $result.variant | Should Be 'original'
        $result.mismatch_count | Should Be 0
        $result.prompt_count | Should BeGreaterThan 0
        $result.prompt_dir | Should Match 'prompts'
    }
}
