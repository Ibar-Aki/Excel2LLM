$projectRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$extractScript = Join-Path $projectRoot 'scripts\extract_excel.ps1'
$packScript = Join-Path $projectRoot 'scripts\pack_for_llm.ps1'
$verifyScript = Join-Path $projectRoot 'scripts\excel_verify.ps1'
$sampleScript = Join-Path $projectRoot 'scripts\create_sample_workbook.ps1'

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
}
