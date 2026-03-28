[CmdletBinding()]
param(
    [string]$OutputPath
)

$ErrorActionPreference = 'Stop'

if (-not $OutputPath) {
    $OutputPath = Join-Path (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'docs\reference') 'TEST_SPECIFICATION.xlsx'
}

function Ensure-Directory {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Convert-CellText {
    param(
        $Value
    )

    if ($null -eq $Value) {
        return ''
    }

    return ([string]$Value) -replace ' \| ', "`n"
}

function Set-WorksheetTable {
    param(
        [Parameter(Mandatory)]
        $Worksheet,
        [Parameter(Mandatory)]
        [object[][]]$Rows,
        [switch]$WrapText,
        [switch]$AutoFilter
    )

    for ($rowIndex = 0; $rowIndex -lt $Rows.Count; $rowIndex++) {
        $row = $Rows[$rowIndex]
        for ($columnIndex = 0; $columnIndex -lt $row.Count; $columnIndex++) {
            $cell = $Worksheet.Cells.Item($rowIndex + 1, $columnIndex + 1)
            $cell.Value2 = Convert-CellText -Value $row[$columnIndex]
            if ($WrapText) {
                $cell.WrapText = $true
            }
            $cell.VerticalAlignment = -4160
        }
    }

    $columnCount = $Rows[0].Count
    $headerRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item(1, $columnCount))
    $allRange = $Worksheet.Range($Worksheet.Cells.Item(1, 1), $Worksheet.Cells.Item($Rows.Count, $columnCount))
    $headerRange.Font.Bold = $true
    $headerRange.Interior.Color = 15773696
    $allRange.Borders.LineStyle = 1
    if ($AutoFilter) {
        $headerRange.AutoFilter() | Out-Null
    }
    $Worksheet.Rows.Item(1).RowHeight = 24
    $Worksheet.Columns.AutoFit() | Out-Null
    [void]$Worksheet.Activate()
    $window = $Worksheet.Application.ActiveWindow
    if ($null -ne $window) {
        $window.SplitRow = 1
        $window.FreezePanes = $true
    }
}

function Convert-ObjectsToRows {
    param(
        [Parameter(Mandatory)]
        [object[]]$Objects,
        [Parameter(Mandatory)]
        [string[]]$Columns
    )

    $rows = [System.Collections.Generic.List[object[]]]::new()
    [void]$rows.Add($Columns)
    foreach ($obj in $Objects) {
        $row = [System.Collections.Generic.List[object]]::new()
        foreach ($column in $Columns) {
            [void]$row.Add($obj.$column)
        }
        [void]$rows.Add($row.ToArray())
    }
    return $rows.ToArray()
}

$overviewRows = @(
    @('項目', '内容'),
    @('文書名', 'Excel2LLM テスト仕様書'),
    @('作成日', '2026-03-28 10:44 JST'),
    @('作成者', 'Codex (GPT-5)'),
    @('対象プロジェクト', 'C:\Work_Codex\Excel2LLM'),
    @('対象ブランチ', 'main'),
    @('参照HEAD', 'e130b8d'),
    @('対象状態', '2026-03-28 時点の未コミット修正を含むワークツリー'),
    @('主対象機能', 'extract / pack / verify / rebuild / prompt bundle / run_all / share package'),
    @('試験レベル', '統合試験 / 入口試験 / 業務シナリオ試験'),
    @('自動統合試験結果', 'Passed: 22 / Failed: 0 / Skipped: 0 / Pending: 0 / Inconclusive: 0'),
    @('自動統合試験所要時間', '319.75s'),
    @('自己診断試験', 'run_self_test.bat 成功'),
    @('配布パッケージ再生成', 'run_build_share_package.bat 成功'),
    @('補助スモーク試験', 'run_all cmd 経路 / PowerShell 経路 / sheet filter / git diff --check 成功'),
    @('総合判定', '主要フローに対する重大または高優先の未修正不具合は今回の試験では未検出')
)

$fixtureRows = @(
    @('Fixture ID', '名称', '生成元', '内容', '主な検証観点'),
    @('FX-001', 'sample.xlsx', 'scripts\create_sample_workbook.ps1', 'Summary / WideTable / Calc の 3 シート。merge, hidden row/column, hyperlink, comment, freeze panes, wide table, formula を含む。', '基本抽出 / 数式 / hidden / merge / pack / verify / rebuild'),
    @('FX-002', 'sample.xlsm', 'scripts\create_sample_workbook.ps1', 'sample.xlsx と同構造の macro-enabled 形式。', 'has_vba 判定 / xlsm 由来 rebuild'),
    @('FX-003', 'security.xlsx', 'tests\TestHelpers.ps1 / New-MiniWorkbook', '最小 Grid ブック。1 式、1 merge。', 'path redaction / macro opt-in / run_all flag 伝播'),
    @('FX-004', 'grid.xlsx', 'tests\TestHelpers.ps1 / New-MiniWorkbook', '3x4 の最小 Grid ブック。', 'chunking 差分'),
    @('FX-005', 'styles.xlsx', 'tests\TestHelpers.ps1 / New-MiniWorkbook -IncludeStyles', 'fill color / font color / wrap text / border を含む最小ブック。', 'styles.json / style pack / style rebuild'),
    @('FX-006', 'verify.xlsx', 'tests\TestHelpers.ps1 / New-MiniWorkbook', '最小ブック抽出後に JSON 改ざん。', 'verify mismatch 検出'),
    @('FX-007', 'mechanical_original.xlsx', 'scripts\create_domain_sample_workbooks.ps1', '機械設計の改善前ブック。Inputs / Calc / Review。', 'prompt bundle / domain acceptance / 計算シナリオ'),
    @('FX-008', 'mechanical_improved.xlsx', 'scripts\create_domain_sample_workbooks.ps1', '機械設計の改善後ブック。Inputs / ShaftSizing / Checks。', '改善版比較 / 説明性'),
    @('FX-009', 'accounting_original.xlsx', 'scripts\create_domain_sample_workbooks.ps1', '会計の改善前ブック。Transactions / Budget / Summary / Notes。', 'prompt bundle / domain acceptance / 予実分析'),
    @('FX-010', 'accounting_improved.xlsx', 'scripts\create_domain_sample_workbooks.ps1', '会計の改善後ブック。Transactions / Budget / Summary / Checks。', '改善版比較 / 利益率 / 差異列'),
    @('FX-011', 'tampered prompt manifest', 'tests\Excel2LLM.Tests.ps1', '外部ファイルパスを書き込んだ prompt_bundle_manifest.json。', 'cleanup ガード'),
    @('FX-012', 'outside share-package dir', 'tests\Excel2LLM.Tests.ps1', 'distribution 配下外の既存ディレクトリ。', 'share package 削除ガード')
)

$testCaseTsv = @'
ID	Name	Feature	Objective	Data	Preconditions	Steps	Expected	Automation	Latest	Evidence
TC-IT-001	bat 入口 usage 表示	run_*.bat	引数不足時に Mandatory prompt に落ちず usage を返すことを確認	bat 6本	bat ファイルが存在	1) 引数なし実行 | 2) -h / --help / /? 実行	終了コード 1 | 使い方 表示 | docs\guides 参照表示	自動	Pass (2026-03-28)	run_tests.bat / Excel2LLM.Tests.ps1
TC-IT-002	基本抽出と xlsm メタデータ	extract	sheet_count, formula, merge, hidden, has_vba の取得を確認	sample.xlsx, sample.xlsm	Excel COM 利用可	1) sample.xlsx 抽出 | 2) workbook.json と manifest.json 確認 | 3) sample.xlsm 抽出	sheet_count=3 | formula_count>3 | Summary merge 1件 | hidden row 7 / hidden col E | Calc!A3 formula/formula2 正常 | .xlsm で has_vba=true	自動	Pass (2026-03-28)	run_tests.bat / workbook.json / manifest.json
TC-IT-003	extract/pack/verify のサマリー表示	console summary	人間向け結果サマリーと次アクション表示を確認	sample.xlsx	bat 利用可	1) run_extract.bat | 2) run_pack.bat | 3) run_verify.bat	抽出/パッキング/検証の各見出し表示 | === 次のおすすめ === 表示	自動	Pass (2026-03-28)	run_tests.bat console output
TC-IT-004	run_all end-to-end	run_all	extract -> pack と verify 付き一括経路を確認	sample.xlsx	run_all.bat 利用可	1) run_all.bat sample.xlsx | 2) run_all.bat sample.xlsx -Verify	workbook.json / llm_package.jsonl 生成 | verify 付きで verify_report.json 生成 | === 一括実行結果 === 表示	自動	Pass (2026-03-28)	run_tests.bat / output files
TC-IT-005	run_all の security flag 伝播	run_all + verify	-RedactPaths と -AllowWorkbookMacros が内部 verify にも届くことを確認	security.xlsx	run_all.bat 利用可	1) run_all.bat security.xlsx -Verify -RedactPaths -AllowWorkbookMacros | 2) workbook.json / verify_report.json 確認	workbook.sheet_count=1 | verify_report.workbook_path=security.xlsx | verify_report.workbook_json_path=workbook.json	自動	Pass (2026-03-28)	run_tests.bat / regression fix
TC-IT-006	run_prompt_bundle wrapper	prompt bundle	wrapper が既定の出力を使って動作することを確認	sample.xlsx 抽出結果	workbook.json と llm_package.jsonl が存在	1) extract | 2) pack | 3) run_prompt_bundle.bat -Scenario general	prompt_bundle_manifest.json 生成 | === 指示文セット作成結果 === 表示 | === 次のおすすめ === 表示	自動	Pass (2026-03-28)	run_tests.bat / prompt_bundle_manifest.json
TC-IT-007	失敗時の番号付き復旧ガイド	error guidance	失敗時の 3 ステップ案内を確認	missing.xlsx	存在しない入力パス	run_all.bat missing.xlsx	非0終了 | 1. Excel を閉じる | 2. コマンドをもう一度実行する | 3. まだダメなら run_self_test.bat	自動	Pass (2026-03-28)	run_tests.bat console output
TC-IT-008	extract/verify の redaction と macro opt-in	security options	RedactPaths と AllowWorkbookMacros の組み合わせ確認	security.xlsx	PowerShell 直実行可	1) extract -RedactPaths -AllowWorkbookMacros | 2) verify -RedactPaths -AllowWorkbookMacros	workbook/manifest/verify_report のパスが basename 化 | verify status success	自動	Pass (2026-03-28)	run_tests.bat / workbook.json / verify_report.json
TC-IT-009	シートフィルタと downstream 整合	extract filter	-Sheets と -ExcludeSheets が pack / verify と整合することを確認	sample.xlsx	filter オプション利用	1) extract -Sheets Summary,Calc,MissingSheet -ExcludeSheets Calc,GhostSheet | 2) pack | 3) verify	抽出シートは Summary のみ | manifest.source_sheet_count=3 | warning に MissingSheet/GhostSheet | verify success	自動	Pass (2026-03-28)	run_tests.bat / manifest.sheet_filter
TC-IT-010	bat 経由のカンマ区切りシート指定	bat + extract filter	-Sheets Summary,Calc の実運用入力を確認	sample.xlsx	run_extract.bat 利用可	run_extract.bat sample.xlsx -Sheets Summary,Calc -ExcludeSheets Calc	Summary のみ抽出 | sheet_filter.include に Summary と Calc | sheet_filter.selected は Summary のみ	自動	Pass (2026-03-28)	run_tests.bat / regression fix
TC-IT-011	sheet と range の chunking 差	pack	chunk mode の違いが chunk range と cell count に反映されることを確認	grid.xlsx	extract 済み	1) pack -ChunkBy sheet -MaxCells 5 | 2) pack -ChunkBy range -MaxCells 5	sheet first chunk=A1:D1, 4 cells | range first chunk=A1:D2, 5 cells	自動	Pass (2026-03-28)	run_tests.bat / jsonl compare
TC-IT-012	style 取得と style 付き pack	extract styles + pack	CollectStyles 時のみ style payload が付くことを確認	styles.xlsx	IncludeStyles 使用	1) extract -CollectStyles | 2) pack -IncludeStyles	style_export_status=generated | styles count > 0 | A1.fill_color=#FF0000 | A1.wrap_text=true	自動	Pass (2026-03-28)	run_tests.bat / styles.json / jsonl
TC-IT-013	tampered workbook.json の mismatch 検出	verify	改ざん後の verify warning を確認	verify.xlsx	extract 済み	1) workbook.json の A1.text を改ざん | 2) verify	verify_report.status=warning | mismatch_count>0 | manifest.verify_status=warning	自動	Pass (2026-03-28)	run_tests.bat / verify_report.json
TC-IT-014	round-trip 完全性	extract + rebuild	主要構造の round-trip を確認	sample.xlsx	rebuild 利用可	1) extract | 2) rebuild | 3) rebuilt xlsx を extract	sheet count / merge count 一致 | freeze panes, hidden row/col 維持 | Calc!A3 formula 維持 | Summary hyperlink/comment 維持 | rebuild_report success	自動	Pass (2026-03-28)	run_tests.bat / roundtrip workbook.json
TC-IT-015	xlsm 由来データの xlsx 復元	rebuild	source_has_vba 付きでも .xlsx 出力になることを確認	sample.xlsm	extract 済み	1) extract sample.xlsm | 2) rebuild | 3) rebuild_report.json 確認	rebuilt file exists | output_extension=.xlsx | source_has_vba=true | warning に VBA 未復元記録	自動	Pass (2026-03-28)	run_tests.bat / rebuild_report.json
TC-IT-016	styles.json あり/なし rebuild 差	rebuild styles	styles.json がある場合だけ見た目復元することを確認	styles.xlsx	style 付き最小ブック	1) extract -CollectStyles | 2) styles.json ありで rebuild | 3) styles.json なしで rebuild | 4) 再抽出して比較	with-style A1.fill_color=#FF0000, wrap_text=true | without-style は同じにならない	自動	Pass (2026-03-28)	run_tests.bat / styles roundtrip
TC-IT-017	workbook.json 欠落時 fail-fast	rebuild invalid input	無効入力で即時失敗することを確認	missing-workbook.json	存在しないパス	rebuild_excel.ps1 -WorkbookJsonPath missing	例外送出 / サイレント成功なし	自動	Pass (2026-03-28)	run_tests.bat
TC-IT-018	機械設計サンプルと prompt bundle	domain prompts mechanical	domain sample -> extract -> pack -> prompt bundle を確認	mechanical_original.xlsx	domain sample generator 利用可	1) domain sample 生成 | 2) extract -CollectStyles | 3) pack -IncludeStyles | 4) prompt bundle を 2 回実行	prompt_count>0 | 先頭 prompt の sheet_name=Calc | keep.txt は残る	自動	Pass (2026-03-28)	run_tests.bat / prompt_bundle_manifest.json
TC-IT-019	prompt manifest の path redaction	prompt bundle security	path 秘匿と cleanup 両立を確認	accounting_original.xlsx	RedactPaths 指定	1) extract -RedactPaths | 2) pack | 3) prompt bundle -RedactPaths を 2 回実行	manifest.workbook_json_path=workbook.json | manifest.jsonl_path=llm_package.jsonl | prompt path が相対または basename	自動	Pass (2026-03-28)	run_tests.bat / prompt_bundle_manifest.json
TC-IT-020	tampered prompt manifest の外部削除防止	prompt cleanup guard	出力先外ファイルが削除されないことを確認	mechanical_original.xlsx, outside.txt	tampered manifest 準備	1) 正常抽出と pack | 2) manifest に外部パスを埋め込む | 3) prompt bundle 再実行	outside.txt が残る | warning に outside output directory	自動	Pass (2026-03-28)	run_tests.bat / prompt cleanup guard
TC-IT-021	share package 削除ガード	share package security	distribution 配下外の無防備削除を防ぐことを確認	outside share-package dir	placeholder.txt を事前配置	1) フラグなしで build_share_package -OutputDir outside | 2) 許可フラグ付きで再実行	1回目は例外 | placeholder.txt は残る | 2回目は package 生成 | share_manifest から絶対パス項目除去 | README/GETTING_STARTED/run_all/run_prompt_bundle 同梱	自動	Pass (2026-03-28)	run_tests.bat / share_manifest.json
TC-IT-022	会計シナリオ domain acceptance	domain acceptance	accounting original で end-to-end 受け入れを確認	accounting_original.xlsx	run_domain_acceptance.ps1 利用可	run_domain_acceptance.ps1 -Scenario accounting -Variant original	result_count=1 | scenario=accounting | variant=original | mismatch_count=0 | prompt_count>0	自動	Pass (2026-03-28)	run_tests.bat / scenario_summary.json
'@ | ConvertFrom-Csv -Delimiter ([char]9)

$executedRunRows = @(
    @('実施ID', '実施日', 'コマンド', '対象', '結果', '所要時間', '確認事項'),
    @('ER-001', '2026-03-28', 'C:\Work_Codex\Excel2LLM\run_tests.bat', '自動統合試験 22 件', 'Passed: 22 / Failed: 0 / Skipped: 0 / Pending: 0 / Inconclusive: 0', '319.75s', '自動統合試験の最新ベースライン'),
    @('ER-002', '2026-03-28', 'C:\Work_Codex\Excel2LLM\run_self_test.bat', '自己診断試験', '成功', '約 50s', 'sample 作成、extract、pack、verify、rebuild、roundtrip extract'),
    @('ER-003', '2026-03-28', 'C:\Work_Codex\Excel2LLM\run_build_share_package.bat', '配布パッケージ再生成', '成功', '短時間', 'distribution\Excel2LLM_Share を再生成'),
    @('ER-004', '2026-03-28', 'run_all.bat sample.xlsx -Verify -RedactPaths -AllowWorkbookMacros', 'cmd 経路スモーク', '成功', '補助確認', 'run_all で verify flag / redaction / macro opt-in が両立'),
    @('ER-005', '2026-03-28', 'run_all.bat sample.xlsx -Verify -Sheets Summary,Calc -ExcludeSheets Calc', 'cmd 経路シートフィルタスモーク', '成功', '補助確認', 'bat 経由のカンマ区切りシート指定を確認'),
    @('ER-006', '2026-03-28', 'git diff --check', '差分健全性確認', '成功', '短時間', '改行・行末空白の問題なし')
)

$coverageRows = @(
    @('区分', '内容', '現状', '補足'),
    @('カバー済み', '抽出の基本品質', '自動化済み', 'sheet_count, formula, merge, hidden, has_vba を検証'),
    @('カバー済み', 'pack の chunking 差分', '自動化済み', 'sheet / range の差を range と cell count で確認'),
    @('カバー済み', 'verify の mismatch 検出', '自動化済み', 'tampered workbook.json で warning を確認'),
    @('カバー済み', 'rebuild round-trip', '自動化済み', 'formula, merge, hyperlink, comment, hidden, freeze panes を確認'),
    @('カバー済み', 'path redaction', '自動化済み', 'extract, verify, prompt bundle で確認'),
    @('カバー済み', 'prompt cleanup guard', '自動化済み', 'manifest 改ざん時の外部削除防止'),
    @('カバー済み', 'share package guard', '自動化済み', '配下外削除禁止と同梱物確認'),
    @('未カバー', '.xls 本格対応', '未実施', '現行スコープ外'),
    @('未カバー', '外部参照ブック', '未実施', 'verify での挙動未固定'),
    @('未カバー', '条件付き書式の完全比較', '未実施', 'styles は best effort'),
    @('未カバー', 'threaded comment 完全復元', '未実施', '通常コメントへのフォールバックが前提'),
    @('未カバー', '名前定義・テーブル・ピボット', '未実施', '現行仕様の正式対象外'),
    @('未カバー', '巨大実ファイルの定点性能試験', '一部のみ', 'wide table はあるが継続ベンチマークは未整備'),
    @('レビュー修正', 'run_all の引数束縛回帰', '修正済み', '2026-03-28 に TC-IT-005, TC-IT-010 で固定'),
    @('レビュー修正', '空 cells 配列での pack 例外', '修正済み', 'AllowEmptyCollection 追加で修正')
)

$excel = $null
$workbook = $null
try {
    Ensure-Directory -Path (Split-Path -Path $OutputPath -Parent)

    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false

    $workbook = $excel.Workbooks.Add()
    while ($workbook.Worksheets.Count -lt 5) {
        [void]$workbook.Worksheets.Add()
    }

    $overviewSheet = $workbook.Worksheets.Item(1)
    $overviewSheet.Name = 'Overview'
    Set-WorksheetTable -Worksheet $overviewSheet -Rows $overviewRows -WrapText

    $fixtureSheet = $workbook.Worksheets.Item(2)
    $fixtureSheet.Name = 'Fixtures'
    Set-WorksheetTable -Worksheet $fixtureSheet -Rows $fixtureRows -WrapText -AutoFilter

    $testCaseSheet = $workbook.Worksheets.Item(3)
    $testCaseSheet.Name = 'TestCases'
    $testCaseRows = Convert-ObjectsToRows -Objects $testCaseTsv -Columns @('ID', 'Name', 'Feature', 'Objective', 'Data', 'Preconditions', 'Steps', 'Expected', 'Automation', 'Latest', 'Evidence')
    Set-WorksheetTable -Worksheet $testCaseSheet -Rows $testCaseRows -WrapText -AutoFilter
    $testCaseSheet.Columns.Item(2).ColumnWidth = 28
    $testCaseSheet.Columns.Item(3).ColumnWidth = 24
    $testCaseSheet.Columns.Item(4).ColumnWidth = 32
    $testCaseSheet.Columns.Item(5).ColumnWidth = 28
    $testCaseSheet.Columns.Item(6).ColumnWidth = 22
    $testCaseSheet.Columns.Item(7).ColumnWidth = 36
    $testCaseSheet.Columns.Item(8).ColumnWidth = 42
    $testCaseSheet.Columns.Item(11).ColumnWidth = 28

    $runSheet = $workbook.Worksheets.Item(4)
    $runSheet.Name = 'ExecutedRuns'
    Set-WorksheetTable -Worksheet $runSheet -Rows $executedRunRows -WrapText -AutoFilter

    $coverageSheet = $workbook.Worksheets.Item(5)
    $coverageSheet.Name = 'CoverageGaps'
    Set-WorksheetTable -Worksheet $coverageSheet -Rows $coverageRows -WrapText -AutoFilter

    foreach ($sheet in @($overviewSheet, $fixtureSheet, $testCaseSheet, $runSheet, $coverageSheet)) {
        $sheet.Columns.Item(1).ColumnWidth = 18
    }

    if (Test-Path -LiteralPath $OutputPath) {
        Remove-Item -LiteralPath $OutputPath -Force
    }
    $workbook.SaveAs($OutputPath, 51)
    Write-Host "Created Excel test specification -> $OutputPath"
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($true) } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($workbook)
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    }
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
