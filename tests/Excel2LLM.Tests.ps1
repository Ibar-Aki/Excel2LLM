$projectRoot = Split-Path -Path $PSScriptRoot -Parent
. (Join-Path $PSScriptRoot 'TestHelpers.ps1')

$extractScript = Join-Path $projectRoot 'scripts\extract_excel.ps1'
$preflightScript = Join-Path $projectRoot 'scripts\preflight_excel.ps1'
$packScript = Join-Path $projectRoot 'scripts\pack_for_llm.ps1'
$verifyScript = Join-Path $projectRoot 'scripts\excel_verify.ps1'
$rebuildScript = Join-Path $projectRoot 'scripts\rebuild_excel.ps1'
$sampleScript = Join-Path $projectRoot 'scripts\create_sample_workbook.ps1'
$domainSampleScript = Join-Path $projectRoot 'scripts\create_domain_sample_workbooks.ps1'
$promptBundleScript = Join-Path $projectRoot 'scripts\export_prompt_bundle.ps1'
$acceptanceScript = Join-Path $projectRoot 'scripts\run_domain_acceptance.ps1'
$sharePackageScript = Join-Path $projectRoot 'scripts\build_share_package.ps1'
$excel2llmBat = Join-Path $projectRoot 'Excel2LLM.bat'
$runExtractBat = Join-Path $projectRoot 'tools\advanced\run_extract.bat'
$runPackBat = Join-Path $projectRoot 'tools\advanced\run_pack.bat'
$runPreflightBat = Join-Path $projectRoot 'tools\advanced\run_preflight.bat'
$runVerifyBat = Join-Path $projectRoot 'tools\advanced\run_verify.bat'
$runRebuildBat = Join-Path $projectRoot 'tools\advanced\run_rebuild.bat'
$runAllBat = Join-Path $projectRoot 'tools\user\run_all.bat'
$runPromptBundleBat = Join-Path $projectRoot 'tools\user\run_prompt_bundle.bat'

Describe 'Excel2LLM integration tests' {
    It 'shows usage help from bat entrypoints instead of falling into PowerShell mandatory prompts' {
        $batCases = @(
            @{ Path = $excel2llmBat; Usage = '使い方: Excel2LLM.bat'; SkipNoArg = $true }
            @{ Path = $runExtractBat; Usage = '使い方: tools\advanced\run_extract.bat' }
            @{ Path = $runPackBat; Usage = '使い方: tools\advanced\run_pack.bat' }
            @{ Path = $runPreflightBat; Usage = '使い方: tools\advanced\run_preflight.bat' }
            @{ Path = $runVerifyBat; Usage = '使い方: tools\advanced\run_verify.bat' }
            @{ Path = $runRebuildBat; Usage = '使い方: tools\advanced\run_rebuild.bat' }
            @{ Path = $runAllBat; Usage = '使い方: tools\user\run_all.bat' }
            @{ Path = $runPromptBundleBat; Usage = '使い方: tools\user\run_prompt_bundle.bat' }
        )

        foreach ($batCase in $batCases) {
            $usagePattern = [regex]::Escape([string]$batCase.Usage)
            if (-not [bool]$batCase['SkipNoArg']) {
                $noArgOutput = & cmd.exe /d /c """$($batCase.Path)""" 2>&1 | Out-String
                $LASTEXITCODE | Should Be 1
                $noArgOutput | Should Match $usagePattern
                $noArgOutput | Should Match 'GETTING_STARTED\.md'
            }

            foreach ($helpSwitch in @('-h', '--help', '/?')) {
                $helpOutput = & cmd.exe /d /c """$($batCase.Path)"" $helpSwitch" 2>&1 | Out-String
                $LASTEXITCODE | Should Be 1
                $helpOutput | Should Match $usagePattern
            }
        }
    }

    It 'writes successful preflight reports for supported workbook types' {
        $workspace = New-TestWorkspace -Name 'preflight-success'
        $samplesDir = Join-Path $workspace 'samples'
        $xlsxOutputDir = Join-Path $workspace 'xlsx-output'
        $xlsmOutputDir = Join-Path $workspace 'xlsm-output'

        & $sampleScript -OutputDir $samplesDir
        & $preflightScript -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $xlsxOutputDir
        & $preflightScript -ExcelPath (Join-Path $samplesDir 'sample.xlsm') -OutputDir $xlsmOutputDir -RedactPaths

        $xlsxReport = Get-Content -LiteralPath (Join-Path $xlsxOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json
        $xlsmReport = Get-Content -LiteralPath (Join-Path $xlsmOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json

        $xlsxReport.status | Should Be 'success'
        $xlsxReport.blocked | Should Be $false
        $xlsxReport.sheet_count | Should Be 3
        $xlsxReport.estimated_total_cells | Should BeGreaterThan 0
        $xlsmReport.status | Should Be 'success'
        $xlsmReport.blocked | Should Be $false
        $xlsmReport.workbook_path | Should Be 'sample.xlsm'
    }

    It 'blocks corrupt and structurally invalid OpenXML workbooks during preflight' {
        $workspace = New-TestWorkspace -Name 'preflight-invalid'
        $corruptPath = Join-Path $workspace 'corrupt.xlsx'
        $missingRelsPath = Join-Path $workspace 'missing-rels.xlsx'
        $corruptOutputDir = Join-Path $workspace 'corrupt-output'
        $missingRelsOutputDir = Join-Path $workspace 'missing-rels-output'

        New-CorruptWorkbookFile -Path $corruptPath
        New-PreflightWorkbookFixture -Path $missingRelsPath -RemoveWorkbookRelationships

        $corruptDidThrow = $false
        try {
            & $preflightScript -ExcelPath $corruptPath -OutputDir $corruptOutputDir
        }
        catch {
            $corruptDidThrow = $true
        }

        $missingRelsDidThrow = $false
        try {
            & $preflightScript -ExcelPath $missingRelsPath -OutputDir $missingRelsOutputDir
        }
        catch {
            $missingRelsDidThrow = $true
        }

        $corruptReport = Get-Content -LiteralPath (Join-Path $corruptOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json
        $missingRelsReport = Get-Content -LiteralPath (Join-Path $missingRelsOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json

        $corruptDidThrow | Should Be $true
        $corruptReport.status | Should Be 'blocked'
        ($corruptReport.reasons -join ' ') | Should Match 'OpenXML ZIP'
        $missingRelsDidThrow | Should Be $true
        $missingRelsReport.status | Should Be 'blocked'
        ($missingRelsReport.reasons -join ' ') | Should Match 'xl/_rels/workbook\.xml\.rels'
    }

    It 'warns on medium-sized workbooks without blocking preflight' {
        $workspace = New-TestWorkspace -Name 'preflight-warning'
        $bookPath = Join-Path $workspace 'warning.xlsx'
        $outputDir = Join-Path $workspace 'output'

        New-PreflightWorkbookFixture -Path $bookPath -PadToBytes 60MB
        & $preflightScript -ExcelPath $bookPath -OutputDir $outputDir

        $preflightReport = Get-Content -LiteralPath (Join-Path $outputDir 'preflight_report.json') -Raw | ConvertFrom-Json

        $preflightReport.status | Should Be 'warning'
        $preflightReport.blocked | Should Be $false
        $preflightReport.file_size_bytes | Should BeGreaterThan 50MB
        ($preflightReport.warnings -join ' ') | Should Match 'ファイルサイズが大きめです'
    }

    It 'blocks oversized or malformed-dimension workbooks before extraction starts' {
        $workspace = New-TestWorkspace -Name 'preflight-blocked-extract'
        $oversizedPath = Join-Path $workspace 'oversized.xlsx'
        $missingDimensionPath = Join-Path $workspace 'missing-dimension.xlsx'
        $oversizedOutputDir = Join-Path $workspace 'oversized-output'
        $missingDimensionOutputDir = Join-Path $workspace 'missing-dimension-output'

        New-PreflightWorkbookFixture -Path $oversizedPath -PadToBytes 205MB
        New-PreflightWorkbookFixture -Path $missingDimensionPath -RemoveWorksheetDimension -PadToBytes 60MB

        $extractOutput = & cmd.exe /d /c """$runExtractBat"" ""$oversizedPath"" -OutputDir ""$oversizedOutputDir""" 2>&1 | Out-String
        $extractExitCode = $LASTEXITCODE
        $missingDimensionDidThrow = $false
        try {
            & $preflightScript -ExcelPath $missingDimensionPath -OutputDir $missingDimensionOutputDir
        }
        catch {
            $missingDimensionDidThrow = $true
        }

        $extractExitCode | Should Not Be 0
        $extractOutput | Should Match '事前チェックで処理を中止しました'
        (Test-Path -LiteralPath (Join-Path $oversizedOutputDir 'workbook.json')) | Should Be $false
        (Test-Path -LiteralPath (Join-Path $oversizedOutputDir 'preflight_report.json')) | Should Be $true

        $oversizedReport = Get-Content -LiteralPath (Join-Path $oversizedOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json
        $missingDimensionReport = Get-Content -LiteralPath (Join-Path $missingDimensionOutputDir 'preflight_report.json') -Raw | ConvertFrom-Json

        $oversizedReport.status | Should Be 'blocked'
        ($oversizedReport.reasons -join ' ') | Should Match 'ファイルサイズが上限を超えています'
        $missingDimensionDidThrow | Should Be $true
        $missingDimensionReport.status | Should Be 'blocked'
        ($missingDimensionReport.reasons -join ' ') | Should Match 'シート範囲情報が見つかりません'
    }

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

    It 'prints human-friendly summaries for extract, pack, and verify commands' {
        $workspace = New-TestWorkspace -Name 'console-summaries'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'
        $jsonlPath = Join-Path $workspace 'llm_package.jsonl'
        $bookPath = Join-Path $samplesDir 'sample.xlsx'
        $workbookJsonPath = Join-Path $outputDir 'workbook.json'

        & $sampleScript -OutputDir $samplesDir

        $extractOutput = & cmd.exe /d /c """$runExtractBat"" ""$bookPath"" -OutputDir ""$outputDir""" 2>&1 | Out-String
        $packOutput = & cmd.exe /d /c """$runPackBat"" ""$workbookJsonPath"" -OutputPath ""$jsonlPath""" 2>&1 | Out-String
        $verifyOutput = & cmd.exe /d /c """$runVerifyBat"" ""$bookPath"" -WorkbookJsonPath ""$workbookJsonPath"" -OutputDir ""$outputDir""" 2>&1 | Out-String

        $extractOutput | Should Match '=== Excel2LLM 抽出結果 ==='
        $extractOutput | Should Match '処理シート:'
        $extractOutput | Should Match '数式数:'
        $extractOutput | Should Match '=== 次のおすすめ ==='
        $packOutput | Should Match '=== パッキング結果 ==='
        $packOutput | Should Match 'チャンク数:'
        $packOutput | Should Match '最大トークン推定:'
        $packOutput | Should Match '=== 次のおすすめ ==='
        $verifyOutput | Should Match '=== 検証結果 ==='
        $verifyOutput | Should Match 'mismatch_count: 0'
        $verifyOutput | Should Match '=== 次のおすすめ ==='
    }

    It 'runs Excel2LLM.bat end-to-end and supports optional verify mode' {
        $workspace = New-TestWorkspace -Name 'run-all'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'
        $bookPath = Join-Path $samplesDir 'sample.xlsx'

        & $sampleScript -OutputDir $samplesDir

        $runAllOutput = & cmd.exe /d /c """$excel2llmBat"" ""$bookPath"" -OutputDir ""$outputDir""" 2>&1 | Out-String
        $runAllVerifyOutput = & cmd.exe /d /c """$excel2llmBat"" ""$bookPath"" -Verify -OutputDir ""$outputDir""" 2>&1 | Out-String

        (Test-Path -LiteralPath (Join-Path $outputDir 'workbook.json')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outputDir 'llm_package.jsonl')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outputDir 'verify_report.json')) | Should Be $true
        $runAllOutput | Should Match '=== 一括実行結果 ==='
        $runAllOutput | Should Match 'LLM に渡す対象:'
        $runAllVerifyOutput | Should Match 'verify 実行: あり'
    }

    It 'propagates verify security flags through Excel2LLM.bat' {
        $workspace = New-TestWorkspace -Name 'run-all-security'
        $bookPath = Join-Path $workspace 'security.xlsx'
        $outputDir = Join-Path $workspace 'output'

        New-MiniWorkbook -Path $bookPath

        & cmd.exe /d /c """$excel2llmBat"" ""$bookPath"" -Verify -RedactPaths -AllowWorkbookMacros -OutputDir ""$outputDir""" | Out-Null

        $LASTEXITCODE | Should Be 0

        $workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json
        $workbookJson.workbook.sheet_count | Should Be 1
        $verifyReport.workbook_path | Should Be 'security.xlsx'
        $verifyReport.workbook_json_path | Should Be 'workbook.json'
    }

    It 'runs prompt bundle wrapper with default output files and shows next steps' {
        $workspace = New-TestWorkspace -Name 'prompt-wrapper'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'
        $bookPath = Join-Path $samplesDir 'sample.xlsx'

        & $sampleScript -OutputDir $samplesDir
        & $extractScript -ExcelPath $bookPath -OutputDir $outputDir
        & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath (Join-Path $outputDir 'llm_package.jsonl')

        Push-Location $workspace
        try {
            $promptOutput = & cmd.exe /d /c """$excel2llmBat"" -PromptBundle -Scenario general -WorkbookJsonPath ""$outputDir\workbook.json"" -JsonlPath ""$outputDir\llm_package.jsonl"" -OutputDir ""$outputDir\prompt_bundle""" 2>&1 | Out-String
        }
        finally {
            Pop-Location
        }

        (Test-Path -LiteralPath (Join-Path $outputDir 'prompt_bundle\prompt_bundle_manifest.json')) | Should Be $true
        $promptOutput | Should Match '=== 指示文セット作成結果 ==='
        $promptOutput | Should Match '=== 次のおすすめ ==='
    }

    It 'stores default extraction output in a timestamped run folder and records the latest run path' {
        $workspace = New-TestWorkspace -Name 'default-output-layout'
        $samplesDir = Join-Path $workspace 'samples'
        $bookPath = Join-Path $samplesDir 'sample.xlsx'
        $latestPointerPath = Join-Path $projectRoot 'output\latest_run.txt'
        $previousPointer = if (Test-Path -LiteralPath $latestPointerPath) { Get-Content -LiteralPath $latestPointerPath -Raw } else { $null }
        $createdRunDir = $null

        try {
            & $sampleScript -OutputDir $samplesDir
            & $extractScript -ExcelPath $bookPath

            $createdRunDir = ((Get-Content -LiteralPath $latestPointerPath -Raw).Trim())
            $createdRunDir | Should Match 'sample_\d{8}-\d{6}$'
            (Test-Path -LiteralPath (Join-Path $createdRunDir 'workbook.json')) | Should Be $true
            (Test-Path -LiteralPath (Join-Path $createdRunDir 'manifest.json')) | Should Be $true
        }
        finally {
            if ($createdRunDir -and (Test-Path -LiteralPath $createdRunDir)) {
                Remove-Item -LiteralPath $createdRunDir -Recurse -Force
            }

            if ($null -eq $previousPointer) {
                if (Test-Path -LiteralPath $latestPointerPath) {
                    Remove-Item -LiteralPath $latestPointerPath -Force
                }
            }
            else {
                Set-Content -LiteralPath $latestPointerPath -Value $previousPointer -NoNewline
            }
        }
    }

    It 'uses the latest run folder by default when creating prompt bundles' {
        $workspace = New-TestWorkspace -Name 'prompt-bundle-latest'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'run-output'
        $bookPath = Join-Path $samplesDir 'sample.xlsx'
        $latestPointerPath = Join-Path $projectRoot 'output\latest_run.txt'
        $previousPointer = if (Test-Path -LiteralPath $latestPointerPath) { Get-Content -LiteralPath $latestPointerPath -Raw } else { $null }

        try {
            & $sampleScript -OutputDir $samplesDir
            & $extractScript -ExcelPath $bookPath -OutputDir $outputDir
            & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath (Join-Path $outputDir 'llm_package.jsonl')

            Push-Location $workspace
            try {
                & cmd.exe /d /c """$excel2llmBat"" -PromptBundle -Scenario general" | Out-Null
            }
            finally {
                Pop-Location
            }

            (Test-Path -LiteralPath (Join-Path $outputDir 'prompt_bundle\prompt_bundle_manifest.json')) | Should Be $true
        }
        finally {
            if ($null -eq $previousPointer) {
                if (Test-Path -LiteralPath $latestPointerPath) {
                    Remove-Item -LiteralPath $latestPointerPath -Force
                }
            }
            else {
                Set-Content -LiteralPath $latestPointerPath -Value $previousPointer -NoNewline
            }
        }
    }

    It 'prints numbered Japanese recovery steps on command failure' {
        $workspace = New-TestWorkspace -Name 'error-guidance'
        $missingPath = Join-Path $workspace 'missing.xlsx'

        $output = & cmd.exe /d /c """$excel2llmBat"" ""$missingPath""" 2>&1 | Out-String

        $LASTEXITCODE | Should Not Be 0
        $output | Should Match '1\. Excel を閉じる'
        $output | Should Match '2\. コマンドをもう一度実行する'
        $output | Should Match '3\. まだダメなら Excel2LLM\.bat -SelfTest'
    }

    It 'supports path redaction and explicit macro opt-in for extract and verify flows' {
        $workspace = New-TestWorkspace -Name 'security-options'
        $bookPath = Join-Path $workspace 'security.xlsx'
        $outputDir = Join-Path $workspace 'output'

        New-MiniWorkbook -Path $bookPath
        & $extractScript -ExcelPath $bookPath -OutputDir $outputDir -RedactPaths -AllowWorkbookMacros
        & $verifyScript -ExcelPath $bookPath -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputDir $outputDir -AllowWorkbookMacros -RedactPaths

        $workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json
        $verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json

        $workbookJson.workbook.path | Should Be 'security.xlsx'
        $manifest.workbook_path | Should Be 'security.xlsx'
        $manifest.output_directory | Should Be 'output'
        $verifyReport.workbook_path | Should Be 'security.xlsx'
        $verifyReport.workbook_json_path | Should Be 'workbook.json'
        $verifyReport.status | Should Be 'success'
    }

    It 'filters extracted sheets and keeps downstream pack and verify consistent' {
        $workspace = New-TestWorkspace -Name 'sheet-filters'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'
        $jsonlPath = Join-Path $workspace 'filtered.jsonl'

        & $sampleScript -OutputDir $samplesDir
        & $extractScript -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -OutputDir $outputDir -Sheets @('Summary', 'Calc', 'MissingSheet') -ExcludeSheets @('Calc', 'GhostSheet')
        & $packScript -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputPath $jsonlPath -ChunkBy sheet -MaxCells 20
        & $verifyScript -ExcelPath (Join-Path $samplesDir 'sample.xlsx') -WorkbookJsonPath (Join-Path $outputDir 'workbook.json') -OutputDir $outputDir

        $workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json
        $verifyReport = Get-Content -LiteralPath (Join-Path $outputDir 'verify_report.json') -Raw | ConvertFrom-Json
        $chunks = @(Get-Content -LiteralPath $jsonlPath | ForEach-Object { $_ | ConvertFrom-Json })
        $chunkSheetNames = @($chunks | Select-Object -ExpandProperty sheet_name | Select-Object -Unique)

        $workbookJson.workbook.sheet_count | Should Be 1
        $workbookJson.sheets[0].sheet_name | Should Be 'Summary'
        $manifest.source_sheet_count | Should Be 3
        $manifest.sheet_count | Should Be 1
        (@($manifest.sheet_filter.include) -contains 'Summary') | Should Be $true
        (@($manifest.sheet_filter.include) -contains 'Calc') | Should Be $true
        (@($manifest.sheet_filter.exclude) -contains 'Calc') | Should Be $true
        (@($manifest.sheet_filter.selected) -contains 'Summary') | Should Be $true
        (@($manifest.sheet_filter.selected) -contains 'Calc') | Should Be $false
        ($manifest.warnings -join ' ') | Should Match 'MissingSheet'
        ($manifest.warnings -join ' ') | Should Match 'GhostSheet'
        $chunkSheetNames.Count | Should Be 1
        $chunkSheetNames[0] | Should Be 'Summary'
        $verifyReport.status | Should Be 'success'
    }

    It 'supports comma-delimited sheet filters from bat entrypoints' {
        $workspace = New-TestWorkspace -Name 'sheet-filters-bat'
        $samplesDir = Join-Path $workspace 'samples'
        $outputDir = Join-Path $workspace 'output'

        & $sampleScript -OutputDir $samplesDir

        & cmd.exe /d /c """$runExtractBat"" ""$samplesDir\sample.xlsx"" -OutputDir ""$outputDir"" -Sheets Summary,Calc -ExcludeSheets Calc" | Out-Null
        $LASTEXITCODE | Should Be 0

        $workbookJson = Get-Content -LiteralPath (Join-Path $outputDir 'workbook.json') -Raw | ConvertFrom-Json
        $manifest = Get-Content -LiteralPath (Join-Path $outputDir 'manifest.json') -Raw | ConvertFrom-Json

        $workbookJson.workbook.sheet_count | Should Be 1
        $workbookJson.sheets[0].sheet_name | Should Be 'Summary'
        (@($manifest.sheet_filter.include) -contains 'Summary') | Should Be $true
        (@($manifest.sheet_filter.include) -contains 'Calc') | Should Be $true
        (@($manifest.sheet_filter.selected) -contains 'Summary') | Should Be $true
        (@($manifest.sheet_filter.selected) -contains 'Calc') | Should Be $false
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

    It 'redacts prompt bundle manifest paths while keeping cleanup functional' {
        $workspace = New-TestWorkspace -Name 'prompt-redaction'
        $samplesDir = Join-Path $workspace 'samples'
        $extractDir = Join-Path $workspace 'extract'
        $promptDir = Join-Path $workspace 'prompts'

        & $domainSampleScript -OutputDir $samplesDir -Scenario accounting -Variant original
        & $extractScript -ExcelPath (Join-Path $samplesDir 'accounting_original.xlsx') -OutputDir $extractDir -RedactPaths
        & $packScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -OutputPath (Join-Path $extractDir 'llm_package.jsonl') -ChunkBy range -MaxCells 120
        & $promptBundleScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario accounting -OutputDir $promptDir -RedactPaths
        & $promptBundleScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario accounting -OutputDir $promptDir -RedactPaths

        $manifest = Get-Content -LiteralPath (Join-Path $promptDir 'prompt_bundle_manifest.json') -Raw | ConvertFrom-Json
        $firstPrompt = $manifest.prompts | Select-Object -First 1

        $manifest.workbook_json_path | Should Be 'workbook.json'
        $manifest.jsonl_path | Should Be 'llm_package.jsonl'
        [System.IO.Path]::IsPathRooted([string]$firstPrompt.path) | Should Be $false
        (Test-Path -LiteralPath (Join-Path $promptDir ([string]$firstPrompt.path))) | Should Be $true
    }

    It 'does not delete files outside the prompt output directory when manifest paths are tampered' {
        $workspace = New-TestWorkspace -Name 'prompt-cleanup-guard'
        $samplesDir = Join-Path $workspace 'samples'
        $extractDir = Join-Path $workspace 'extract'
        $promptDir = Join-Path $workspace 'prompts'
        $outsideFile = Join-Path $workspace 'outside.txt'

        & $domainSampleScript -OutputDir $samplesDir -Scenario mechanical -Variant original
        & $extractScript -ExcelPath (Join-Path $samplesDir 'mechanical_original.xlsx') -OutputDir $extractDir
        & $packScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -OutputPath (Join-Path $extractDir 'llm_package.jsonl') -ChunkBy range -MaxCells 120

        Ensure-Directory -Path $promptDir
        Set-Content -LiteralPath $outsideFile -Value 'keep me'
        Write-JsonFile -Path (Join-Path $promptDir 'prompt_bundle_manifest.json') -Data ([ordered]@{
            prompts = @(
                [ordered]@{
                    path = $outsideFile
                }
            )
        })

        & $promptBundleScript -WorkbookJsonPath (Join-Path $extractDir 'workbook.json') -JsonlPath (Join-Path $extractDir 'llm_package.jsonl') -Scenario mechanical -OutputDir $promptDir

        $manifest = Get-Content -LiteralPath (Join-Path $promptDir 'prompt_bundle_manifest.json') -Raw | ConvertFrom-Json
        (Test-Path -LiteralPath $outsideFile) | Should Be $true
        ($manifest.warnings -join ' ') | Should Match 'outside output directory'
    }

    It 'blocks share package cleanup outside distribution unless explicit override flags are provided' {
        $workspace = New-TestWorkspace -Name 'share-package-guard'
        $outsideDir = Join-Path $workspace 'share-package'
        Ensure-Directory -Path $outsideDir
        Set-Content -LiteralPath (Join-Path $outsideDir 'placeholder.txt') -Value 'placeholder'

        $didThrow = $false
        try {
            & $sharePackageScript -OutputDir $outsideDir
        }
        catch {
            $didThrow = $true
        }

        $didThrow | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'placeholder.txt')) | Should Be $true

        & $sharePackageScript -OutputDir $outsideDir -AllowOutsideDistribution -ForceCleanOutputDir

        $shareManifest = Get-Content -LiteralPath (Join-Path $outsideDir 'share_manifest.json') -Raw | ConvertFrom-Json
        (@($shareManifest.PSObject.Properties.Name) -contains 'source_project_root') | Should Be $false
        (@($shareManifest.PSObject.Properties.Name) -contains 'output_directory') | Should Be $false
        $shareManifest.package_name | Should Be 'share-package'
        (Test-Path -LiteralPath (Join-Path $outsideDir 'README.md')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'GETTING_STARTED.md')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'Excel2LLM.bat')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'tools\user\run_all.bat')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'tools\advanced\run_preflight.bat')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'tools\user\run_prompt_bundle.bat')) | Should Be $true
        (Test-Path -LiteralPath (Join-Path $outsideDir 'docs\guides\SHARE_PACKAGE.md')) | Should Be $false
        (Test-Path -LiteralPath (Join-Path $outsideDir 'docs\guides\MANUAL.md')) | Should Be $false
        (Test-Path -LiteralPath (Join-Path $outsideDir 'docs\guides\USER_GUIDE.md')) | Should Be $false
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
