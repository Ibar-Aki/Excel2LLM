# Excel2LLM テスト仕様書

- 作成日: 2026-03-28 10:44 JST
- 作成者: Codex (GPT-5)

## 1. 文書の目的

本書は、`Excel2LLM` の試験範囲、試験方針、試験データ、試験ケース、実施済み試験結果、現時点の既知ギャップを詳細に記録するための仕様書です。

本書は要約版ではなく、次の観点を網羅的に残すことを目的としています。

- 何を試験対象としているか
- どの入力データで試験しているか
- 各試験でどの手順を踏み、何を合格条件にしているか
- どの試験が自動化済みか
- 2026-03-28 時点で実際に何を実行し、どういう結果になったか
- どこまでカバーできていて、どこが未カバーか

## 2. 対象システム

### 2.1 システム名

- `Excel2LLM`

### 2.2 対象ワークツリー

- プロジェクトルート: `C:\Work_Codex\Excel2LLM`
- Git ブランチ: `main`
- 参照 HEAD: `e130b8d`
- 備考: 本書は 2026-03-28 時点の未コミット修正を含むワークツリーを対象とする

### 2.3 主な構成要素

- `scripts\extract_excel.ps1`
  - Excel ブックを `workbook.json` / `styles.json` / `manifest.json` に抽出
- `scripts\pack_for_llm.ps1`
  - `workbook.json` を `llm_package.jsonl` に分割
- `scripts\excel_verify.ps1`
  - 抽出結果と Excel 再計算結果の差分検証
- `scripts\rebuild_excel.ps1`
  - `workbook.json` から `.xlsx` を逆生成
- `scripts\export_prompt_bundle.ps1`
  - `llm_package.jsonl` から LLM 向け prompt 束を生成
- `scripts\run_all.ps1`
  - `extract -> verify(optional) -> pack` を一括実行
- `scripts\build_share_package.ps1`
  - 配布用フォルダを再生成
- `run_*.bat`
  - 利用者向けの実行入口

### 2.4 主な出力

- `workbook.json`
- `styles.json`
- `manifest.json`
- `llm_package.jsonl`
- `verify_report.json`
- `rebuild_report.json`
- `prompt_bundle_manifest.json`
- `distribution\Excel2LLM_Share\*`

## 3. 試験方針

### 3.1 試験レベル

本システムでは次の 3 層を実施対象とします。

1. スクリプト統合試験
   - PowerShell スクリプト単位の主要フローを実行し、出力ファイルと内容を検証する
2. 利用者入口試験
   - `.bat` 経由の使い方、ヘルプ、引数受け渡し、導線表示を検証する
3. 業務シナリオ試験
   - 機械設計、会計のサンプル Excel を用いた end-to-end の受け入れ確認を行う

### 3.2 試験の主眼

- 情報欠落がないこと
- 数式、表示値、結合セル、hidden 状態、freeze panes が保持されること
- `extract -> pack -> verify -> rebuild -> re-extract` の round-trip が成立すること
- パス秘匿、マクロ無効化、削除ガードなどの安全性が崩れていないこと
- `.bat` 経由の利用者導線が壊れていないこと
- 配布用フォルダに必要物が含まれること

### 3.3 非対象

現時点で正式にカバー対象としていないものは次です。

- `.xls` 形式の本格対応
- VBA 本体の復元
- 条件付き書式の完全再現
- 外部リンク参照先の完全再現
- threaded comment の完全復元
- GUI や専用デスクトップ画面

## 4. 試験環境

### 4.1 実行環境

- OS: Windows
- Excel: M365 Excel
- 実行基盤:
  - `powershell.exe`
  - `pwsh`
  - Excel COM
  - `.bat` エントリポイント

### 4.2 前提条件

- Excel COM 自動化が利用可能であること
- ローカルファイルへの読み書き権限があること
- `run_tests.bat` 実行時に Pester が利用可能であること
- サンプル作成用に Excel を起動できること

## 5. 試験データ設計

### 5.1 共通サンプル `sample.xlsx` / `sample.xlsm`

`scripts\create_sample_workbook.ps1` が生成するサンプルで、以下を含みます。

- `Summary` シート
  - 数値セル
  - 数式セル
  - hidden row
  - hidden column
  - merge
  - hyperlink
  - legacy comment
  - freeze panes
- `WideTable` シート
  - 50 行 x 100 列級の wide table
  - 末尾列の数式
- `Calc` シート
  - 単純数式
  - 日付文字列
  - `TEXT(TODAY(), ...)` 系数式
- `sample.xlsm`
  - `has_vba` 判定用の macro-enabled 形式

### 5.2 最小サンプル `New-MiniWorkbook`

`tests\TestHelpers.ps1` が生成する最小ブックで、以下を含みます。

- 3 行 x 4 列の `Grid` シート
- `D3` の数式
- `A2:B2` の merge
- style 付き変種では以下を追加
  - fill color
  - font color
  - wrap text
  - border

### 5.3 ドメインサンプル

`scripts\create_domain_sample_workbooks.ps1` が生成する業務サンプルです。

- 機械設計
  - `mechanical_original.xlsx`
  - `mechanical_improved.xlsx`
- 会計
  - `accounting_original.xlsx`
  - `accounting_improved.xlsx`

含める観点:

- 意味のある数式
- 改善前後の比較
- チェックシート
- サマリー表
- prompt bundle の材料となる説明性

## 6. 自動化済み試験ケース一覧

### 6.1 TC-IT-001

- 試験ID: `TC-IT-001`
- 試験名: `.bat` 入口の usage 表示確認
- 目的:
  - 引数不足時に PowerShell の Mandatory 入力待ちへ落ちず、日本語の usage を返すことを確認する
- 対象:
  - `run_extract.bat`
  - `run_pack.bat`
  - `run_verify.bat`
  - `run_rebuild.bat`
  - `run_all.bat`
  - `run_prompt_bundle.bat`
- 前提条件:
  - `.bat` ファイルが存在すること
- 手順:
  1. 各 `.bat` を引数なしで実行する
  2. 各 `.bat` を `-h`、`--help`、`/?` 付きで実行する
- 期待結果:
  - 終了コードが `1`
  - `Usage: run_xxx.bat` が表示される
  - `docs\guides\` への参照が表示される
- 自動化状態:
  - 自動化済み
- 実装箇所:
  - `tests\Excel2LLM.Tests.ps1`

### 6.2 TC-IT-002

- 試験ID: `TC-IT-002`
- 試験名: 基本抽出と `.xlsm` メタデータ確認
- 目的:
  - `extract` が workbook metadata、数式、merge、hidden row/column、VBA 有無を正しく出すことを確認する
- 使用データ:
  - `sample.xlsx`
  - `sample.xlsm`
- 手順:
  1. `sample.xlsx` を抽出する
  2. `workbook.json` と `manifest.json` を読む
  3. `sample.xlsm` を抽出し直し、`has_vba` を確認する
- 期待結果:
  - `sheet_count = 3`
  - `formula_count > 3`
  - `Summary` に merge が 1 件ある
  - hidden row `7` と hidden column `E` が取得される
  - `Calc!A3.formula = '=SUM(A1:A2)'`
  - `Calc!A3.formula2 = '=SUM(A1:A2)'`
  - `.xlsm` 抽出時に `workbook.has_vba = true`
- 自動化状態:
  - 自動化済み

### 6.3 TC-IT-003

- 試験ID: `TC-IT-003`
- 試験名: `extract` / `pack` / `verify` の人間向けサマリー表示
- 目的:
  - コンソール上に必要な進捗サマリーと「次のおすすめ」が出ることを確認する
- 使用データ:
  - `sample.xlsx`
- 手順:
  1. `run_extract.bat`
  2. `run_pack.bat`
  3. `run_verify.bat`
- 期待結果:
  - `=== Excel2LLM 抽出結果 ===` が出る
  - `=== パッキング結果 ===` が出る
  - `=== 検証結果 ===` が出る
  - 各コマンドで `=== 次のおすすめ ===` が出る
- 自動化状態:
  - 自動化済み

### 6.4 TC-IT-004

- 試験ID: `TC-IT-004`
- 試験名: `run_all.bat` の end-to-end 実行
- 目的:
  - `run_all` が既定で `extract -> pack`、`-Verify` 指定で `extract -> verify -> pack` を完走することを確認する
- 使用データ:
  - `sample.xlsx`
- 手順:
  1. `run_all.bat sample.xlsx -OutputDir ...`
  2. `run_all.bat sample.xlsx -Verify -OutputDir ...`
- 期待結果:
  - `workbook.json` が生成される
  - `llm_package.jsonl` が生成される
  - `-Verify` 時は `verify_report.json` も生成される
  - `=== 一括実行結果 ===` が表示される
- 自動化状態:
  - 自動化済み

### 6.5 TC-IT-005

- 試験ID: `TC-IT-005`
- 試験名: `run_all` の security flag 伝播
- 目的:
  - `run_all` に与えた `-RedactPaths` と `-AllowWorkbookMacros` が、内部の `extract` と `verify` に正しく伝播することを確認する
- 使用データ:
  - `security.xlsx`
- 手順:
  1. `run_all.bat security.xlsx -Verify -RedactPaths -AllowWorkbookMacros`
  2. `workbook.json` と `verify_report.json` を確認する
- 期待結果:
  - `workbook.sheet_count = 1`
  - `verify_report.workbook_path = security.xlsx`
  - `verify_report.workbook_json_path = workbook.json`
- 自動化状態:
  - 自動化済み
- 備考:
  - 2026-03-28 のレビューで発見した回帰を固定するために追加

### 6.6 TC-IT-006

- 試験ID: `TC-IT-006`
- 試験名: `run_prompt_bundle.bat` の既定動作
- 目的:
  - `prompt bundle` の薄いラッパーが既定の出力を使って実行できることを確認する
- 使用データ:
  - `sample.xlsx` 抽出結果
- 手順:
  1. `extract`
  2. `pack`
  3. `run_prompt_bundle.bat -Scenario general ...`
- 期待結果:
  - `prompt_bundle_manifest.json` が生成される
  - `=== Prompt Bundle 結果 ===` が表示される
  - `=== 次のおすすめ ===` が表示される
- 自動化状態:
  - 自動化済み

### 6.7 TC-IT-007

- 試験ID: `TC-IT-007`
- 試験名: 失敗時の番号付き復旧ガイド
- 目的:
  - 実行失敗時にユーザー向けの 3 ステップ対処案内が出ることを確認する
- 使用データ:
  - 存在しない Excel パス
- 手順:
  1. `run_all.bat missing.xlsx`
- 期待結果:
  - 非 0 終了
  - `1. Excel を閉じる`
  - `2. コマンドをもう一度実行する`
  - `3. まだダメなら run_self_test.bat`
    が表示される
- 自動化状態:
  - 自動化済み

### 6.8 TC-IT-008

- 試験ID: `TC-IT-008`
- 試験名: `extract` / `verify` の path redaction と macro opt-in
- 目的:
  - `-RedactPaths` と `-AllowWorkbookMacros` の組み合わせが `extract` と `verify` で成立することを確認する
- 使用データ:
  - `security.xlsx`
- 手順:
  1. `extract_excel.ps1 -RedactPaths -AllowWorkbookMacros`
  2. `excel_verify.ps1 -RedactPaths -AllowWorkbookMacros`
- 期待結果:
  - `workbook.path = security.xlsx`
  - `manifest.workbook_path = security.xlsx`
  - `manifest.output_directory = output`
  - `verify_report.workbook_path = security.xlsx`
  - `verify_report.workbook_json_path = workbook.json`
- 自動化状態:
  - 自動化済み

### 6.9 TC-IT-009

- 試験ID: `TC-IT-009`
- 試験名: シートフィルタと downstream 整合性
- 目的:
  - `-Sheets` / `-ExcludeSheets` が `extract` で機能し、その後の `pack` / `verify` も filtered `workbook.json` に追従することを確認する
- 使用データ:
  - `sample.xlsx`
- 手順:
  1. `extract` を `-Sheets Summary,Calc,MissingSheet -ExcludeSheets Calc,GhostSheet` 付きで実行する
  2. filtered `workbook.json` を `pack`
  3. 同じ filtered `workbook.json` を `verify`
- 期待結果:
  - 抽出結果シート数は `1`
  - 選択シートは `Summary`
  - `manifest.source_sheet_count = 3`
  - `manifest.sheet_filter.include/exclude/selected` が期待どおり
  - warning に `MissingSheet` と `GhostSheet` が残る
  - `verify_report.status = success`
- 自動化状態:
  - 自動化済み

### 6.10 TC-IT-010

- 試験ID: `TC-IT-010`
- 試験名: `.bat` 入口でのカンマ区切りシート指定
- 目的:
  - `run_extract.bat -Sheets Summary,Calc -ExcludeSheets Calc` のような実運用入力が正しく分解されることを確認する
- 使用データ:
  - `sample.xlsx`
- 手順:
  1. `.bat` 経由でカンマ区切りシート指定を付けて抽出する
  2. `workbook.json` と `manifest.json` を確認する
- 期待結果:
  - 抽出結果は `Summary` のみ
  - `sheet_filter.include` に `Summary` と `Calc` が含まれる
  - `sheet_filter.selected` に `Summary` のみが含まれる
- 自動化状態:
  - 自動化済み
- 備考:
  - 2026-03-28 のレビューで発見した `.bat` 実運用経路の回帰を固定するために追加

### 6.11 TC-IT-011

- 試験ID: `TC-IT-011`
- 試験名: `sheet` chunking と `range` chunking の差
- 目的:
  - chunking モード差が実際の `range` と `cell_addresses` に反映されることを確認する
- 使用データ:
  - `grid.xlsx`
- 手順:
  1. `extract`
  2. `pack -ChunkBy sheet -MaxCells 5`
  3. `pack -ChunkBy range -MaxCells 5`
- 期待結果:
  - `sheet` の先頭 chunk は `A1:D1`、セル数 4
  - `range` の先頭 chunk は `A1:D2`、セル数 5
- 自動化状態:
  - 自動化済み

### 6.12 TC-IT-012

- 試験ID: `TC-IT-012`
- 試験名: style 取得と style 付き pack
- 目的:
  - `CollectStyles` 時だけ `styles.json` が充実し、`IncludeStyles` 付き pack で style payload が付くことを確認する
- 使用データ:
  - `styles.xlsx`
- 手順:
  1. style 付き最小ブックを生成
  2. `extract -CollectStyles`
  3. `pack -IncludeStyles`
- 期待結果:
  - `manifest.style_export_status = generated`
  - `styles.json.styles.Count > 0`
  - `A1.style.fill_color = #FF0000`
  - `A1.style.wrap_text = true`
- 自動化状態:
  - 自動化済み

### 6.13 TC-IT-013

- 試験ID: `TC-IT-013`
- 試験名: tampered `workbook.json` の差分検出
- 目的:
  - `verify` が抽出後に改ざんされた JSON を検知できることを確認する
- 使用データ:
  - `verify.xlsx`
- 手順:
  1. `extract`
  2. `workbook.json` の `A1.text` を `BROKEN` に改ざん
  3. `verify`
- 期待結果:
  - `verify_report.status = warning`
  - `verify_report.mismatch_count > 0`
  - `manifest.verify_status = warning`
- 自動化状態:
  - 自動化済み

### 6.14 TC-IT-014

- 試験ID: `TC-IT-014`
- 試験名: round-trip 復元の完全性
- 目的:
  - `extract -> rebuild -> extract` の round-trip で、主要構造と重要情報が保持されることを確認する
- 使用データ:
  - `sample.xlsx`
- 手順:
  1. `extract`
  2. `rebuild`
  3. rebuilt `.xlsx` を再度 `extract`
- 期待結果:
  - sheet 数が一致
  - merge 件数が一致
  - `Summary.freeze_panes` が保持される
  - hidden row / hidden column が保持される
  - `Calc!A3.formula = '=SUM(A1:A2)'`
  - hyperlink と comment が保持される
  - `rebuild_report.status = success`
- 自動化状態:
  - 自動化済み

### 6.15 TC-IT-015

- 試験ID: `TC-IT-015`
- 試験名: `.xlsm` 由来データの `.xlsx` 復元
- 目的:
  - macro-enabled workbook 由来でも、逆生成は `.xlsx` で保存され、VBA 未復元が report に残ることを確認する
- 使用データ:
  - `sample.xlsm`
- 手順:
  1. `extract`
  2. `rebuild`
  3. `rebuild_report.json` を確認
- 期待結果:
  - rebuilt ファイルが存在
  - 出力拡張子は `.xlsx`
  - `source_has_vba = true`
  - warning に VBA 未復元が記録される
- 自動化状態:
  - 自動化済み

### 6.16 TC-IT-016

- 試験ID: `TC-IT-016`
- 試験名: `styles.json` あり/なしでの rebuild 差
- 目的:
  - style 情報がある場合だけ見た目復元が行われることを確認する
- 使用データ:
  - style 付き `styles.xlsx`
- 手順:
  1. `extract -CollectStyles`
  2. `styles.json` ありで `rebuild`
  3. `styles.json` なしで `rebuild`
  4. 両者を再抽出して style を比較
- 期待結果:
  - with-style 側の `A1.fill_color = #FF0000`
  - with-style 側の `A1.wrap_text = true`
  - without-style 側は同値にならない
- 自動化状態:
  - 自動化済み

### 6.17 TC-IT-017

- 試験ID: `TC-IT-017`
- 試験名: `workbook.json` 欠落時の fail-fast
- 目的:
  - invalid input に対して `rebuild` が即時失敗することを確認する
- 使用データ:
  - 存在しない `workbook.json`
- 手順:
  1. `rebuild_excel.ps1` に存在しないパスを渡す
- 期待結果:
  - 例外送出
  - サイレント成功しない
- 自動化状態:
  - 自動化済み

### 6.18 TC-IT-018

- 試験ID: `TC-IT-018`
- 試験名: 機械設計サンプルと prompt bundle 生成
- 目的:
  - domain サンプル生成、抽出、pack、scenario-specific prompt bundle がつながることを確認する
- 使用データ:
  - `mechanical_original.xlsx`
- 手順:
  1. ドメインサンプル生成
  2. `extract -CollectStyles`
  3. `pack -ChunkBy range -IncludeStyles`
  4. `prompt bundle` を 2 回実行して再生成も確認
- 期待結果:
  - sample ファイルが生成される
  - `manifest.scenario = mechanical`
  - prompt 数 > 0
  - 先頭 prompt のシートが `Calc`
  - 無関係な `keep.txt` は削除されない
- 自動化状態:
  - 自動化済み

### 6.19 TC-IT-019

- 試験ID: `TC-IT-019`
- 試験名: prompt manifest の path redaction
- 目的:
  - `prompt bundle` の manifest で path が秘匿されても、cleanup が壊れないことを確認する
- 使用データ:
  - `accounting_original.xlsx`
- 手順:
  1. `extract -RedactPaths`
  2. `pack`
  3. `prompt bundle -RedactPaths` を 2 回実行
- 期待結果:
  - `manifest.workbook_json_path = workbook.json`
  - `manifest.jsonl_path = llm_package.jsonl`
  - prompt path が絶対パスでない
  - 実ファイルは存在する
- 自動化状態:
  - 自動化済み

### 6.20 TC-IT-020

- 試験ID: `TC-IT-020`
- 試験名: tampered prompt manifest による外部削除防止
- 目的:
  - `prompt_bundle_manifest.json` を改ざんしても、出力先外のファイルが削除されないことを確認する
- 使用データ:
  - `mechanical_original.xlsx`
  - prompt 出力先外の `outside.txt`
- 手順:
  1. 正常な抽出と pack を行う
  2. prompt manifest を改ざんし、外部ファイルを `path` に書き込む
  3. `prompt bundle` を再実行する
- 期待結果:
  - `outside.txt` が残る
  - warning に `outside output directory` が記録される
- 自動化状態:
  - 自動化済み

### 6.21 TC-IT-021

- 試験ID: `TC-IT-021`
- 試験名: share package の削除ガード
- 目的:
  - `distribution` 配下外の既存ディレクトリを、明示フラグなしで削除しないことを確認する
- 使用データ:
  - 一時ディレクトリ `share-package`
- 手順:
  1. フラグなしで `build_share_package.ps1 -OutputDir outside`
  2. `-AllowOutsideDistribution -ForceCleanOutputDir` 付きで再実行
- 期待結果:
  - フラグなしでは例外
  - `placeholder.txt` は消えない
  - フラグありでは package が生成される
  - `share_manifest.json` から `source_project_root` / `output_directory` が除去されている
  - `README.md`、`GETTING_STARTED.md`、`run_all.bat`、`run_prompt_bundle.bat` が含まれる
- 自動化状態:
  - 自動化済み

### 6.22 TC-IT-022

- 試験ID: `TC-IT-022`
- 試験名: 会計シナリオの domain acceptance
- 目的:
  - ドメイン受け入れワークフローが accounting original サンプルで成功することを確認する
- 使用データ:
  - `accounting_original.xlsx`
- 手順:
  1. `run_domain_acceptance.ps1 -Scenario accounting -Variant original`
  2. `scenario_summary.json` を確認
- 期待結果:
  - `result_count = 1`
  - `scenario = accounting`
  - `variant = original`
  - `mismatch_count = 0`
  - `prompt_count > 0`
  - `prompt_dir` に `prompts` が含まれる
- 自動化状態:
  - 自動化済み

## 7. 実施済み試験ログ

### 7.1 自動統合試験

- 実施日:
  - 2026-03-28
- 実行コマンド:
  - `C:\Work_Codex\Excel2LLM\run_tests.bat`
- 実行結果:
  - `Passed: 22 Failed: 0 Skipped: 0 Pending: 0 Inconclusive: 0`
- 所要時間:
  - `319.75s`
- 備考:
  - `TC-IT-001` から `TC-IT-022` までを実行

### 7.2 自己診断試験

- 実施日:
  - 2026-03-28
- 実行コマンド:
  - `C:\Work_Codex\Excel2LLM\run_self_test.bat`
- 実行内容:
  - sample workbook 作成
  - extract
  - pack
  - verify
  - rebuild
  - rebuilt workbook の再抽出
- 実行結果:
  - `Self-test completed successfully.`

### 7.3 配布用パッケージ再生成試験

- 実施日:
  - 2026-03-28
- 実行コマンド:
  - `C:\Work_Codex\Excel2LLM\run_build_share_package.bat`
- 実行結果:
  - `Built share package -> C:\Work_Codex\Excel2LLM\distribution\Excel2LLM_Share`

### 7.4 補助スモーク試験

- 実施日:
  - 2026-03-28
- 実行内容:
  - `run_all.bat sample.xlsx -Verify -RedactPaths -AllowWorkbookMacros` を `cmd` 経由で実行
  - 同コマンドを PowerShell から `.bat` 呼び出しで実行
  - `run_all.bat ... -Sheets Summary,Calc -ExcludeSheets Calc` を確認
  - `git diff --check` による差分健全性確認
- 実行結果:
  - すべて成功
  - `git diff --check` 通過

## 8. 2026-03-28 のレビューで発見し修正した内容

### 8.1 `run_all` の可変引数転送回帰

- 症状:
  - `run_all -Verify -RedactPaths -AllowWorkbookMacros` 実行時、`-RedactPaths` や `-AllowWorkbookMacros` が `extract_excel.ps1` 側で `Sheets` / `ExcludeSheets` に誤束縛され、抽出結果が 0 シート・0 セルになるケースがあった
- 影響:
  - 空の `workbook.json`
  - 空の `llm_package.jsonl`
  - 見かけ上は一括実行成功
- 修正:
  - `run_all.ps1` で可変引数転送を廃止し、抽出向けパラメータを明示ハッシュで渡すよう変更
- 固定した試験:
  - `TC-IT-005`

### 8.2 空配列時の `pack` 例外

- 症状:
  - `Group-CellsBySheet` が空配列を受けると例外になり、フィルタ後に 0 セルとなるケースで `pack` が崩れる可能性があった
- 修正:
  - `Group-CellsBySheet` の `Cells` パラメータに `AllowEmptyCollection` を付与
- 固定した試験:
  - `TC-IT-005`

### 8.3 `.bat` 実運用形式のシート指定

- 症状:
  - `-Sheets Summary,Calc` が 1 文字列として扱われ、利用者向け help に書かれている使用例どおりに動かない経路があった
- 修正:
  - `extract_excel.ps1` のシート名正規化でカンマ区切りを分解
- 固定した試験:
  - `TC-IT-010`

## 9. 機能トレーサビリティ

### 9.1 主機能と対応試験

- `.bat` 入口と usage
  - `TC-IT-001`
- 抽出の基本品質
  - `TC-IT-002`
  - `TC-IT-008`
  - `TC-IT-009`
  - `TC-IT-010`
- console サマリーと導線
  - `TC-IT-003`
  - `TC-IT-004`
  - `TC-IT-006`
  - `TC-IT-007`
- `run_all` 導線
  - `TC-IT-004`
  - `TC-IT-005`
- pack ロジック
  - `TC-IT-011`
  - `TC-IT-012`
- verify ロジック
  - `TC-IT-008`
  - `TC-IT-013`
- rebuild ロジック
  - `TC-IT-014`
  - `TC-IT-015`
  - `TC-IT-016`
  - `TC-IT-017`
- prompt bundle
  - `TC-IT-006`
  - `TC-IT-018`
  - `TC-IT-019`
  - `TC-IT-020`
- 配布パッケージ
  - `TC-IT-021`
- 業務受け入れ
  - `TC-IT-022`

## 10. 未カバー領域と今後の試験候補

### 10.1 未カバー

- `.xls` 形式の抽出/再構築
- 外部参照を含むブックの `verify`
- 条件付き書式の見た目差分
- threaded comment の完全 round-trip
- 名前定義、テーブル、ピボットテーブル
- 外部リンクの復元完全性
- 大規模実ファイルでの長時間性能測定の定期化

### 10.2 今後追加価値が高い試験

- `5000+` セルを超える実ファイルでの定点性能試験
- `run_all` と `run_prompt_bundle` の異常系メッセージ詳細試験
- style 情報欠落時の rebuild warning 詳細試験
- `distribution\Excel2LLM_Share` 単体を別マシンへコピーした前提の配布受け入れ試験

## 11. 判定

2026-03-28 時点のワークツリーに対する判定は次のとおりです。

- 自動統合試験:
  - 合格
- 自己診断試験:
  - 合格
- 配布パッケージ再生成:
  - 合格
- 差分健全性:
  - 合格

現時点で、主要フローに対する重大または高優先の未修正不具合は本試験では確認されていません。
