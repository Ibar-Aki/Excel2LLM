# Excel2LLM

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-12

Windows と M365 Excel を前提に、Excel ブックを LLM 向けの正規化 JSON と JSONL に変換するツール群です。追加インストールは前提にせず、`PowerShell`、`Excel COM`、`VBA`、`bat` だけで動作します。

## ディレクトリ構成

- `scripts/`: 抽出、パック、検証、逆生成、サンプル生成の PowerShell スクリプト
- `output/`: 生成される `workbook.json`、`styles.json`、`llm_package.jsonl` などの出力先
- `docs/`: 利用者向けガイド、参照資料、保守者向け手順、検証レポート
- `distribution/`: 他の人へ渡すために再生成できる配布用フォルダ
- `templates/`: 任意で Excel に取り込める VBA テンプレート
- `tests/`: Pester を使った回帰テストとテスト補助

## まず読むドキュメント

- `docs/README.md`: 文書の全体案内。どの用途で何を読むかをここに集約
- `docs/guides/MANUAL.md`: 初めて使う人向けの最優先マニュアル
- `docs/guides/SHARE_PACKAGE.md`: 他の人へ渡すための配布用フォルダの作り方と使い方
- `docs/guides/USER_GUIDE.md`: 最初のセットアップ、実行手順、出力の見方、トラブル対応
- `docs/reference/LLM_PROMPT_FORMATS.md`: Excel2LLM の出力を LLM にどう指示するかの用途別テンプレート集
- `docs/guides/USE_CASES.md`: 実務での使い方、LLM への渡し方、用途別の具体例
- `docs/reports/DOMAIN_SCENARIO_REPORT_20260311.md`: 機械設計向け、会計向けの一気通貫シナリオ検証レポート
- `docs/reports/SECURITY_FILE_OPS_REVIEW_20260312.md`: セキュリティとファイル操作のレビュー結果
- `docs/maintainers/OPERATIONS.md`: テスト、運用、リリース、Git 管理の実務手順
- `docs/reference/FORMAT.md`: JSON と JSONL の構造
- `docs/reference/VBA_HELPER.md`: VBA 補助の使い方

## できること

- 複数シートのセル内容、数式、表示値、結合セル情報を `workbook.json` に正規化保存
- 色、罫線、配置などの見た目情報を `styles.json` に分離保存
- LLM にそのまま流し込みやすい `llm_package.jsonl` を生成
- Excel 再計算後の差分確認を `verify_report.json` と `manifest.json` に保存
- `workbook.json` から `.xlsx` を逆生成し、`rebuild_report.json` に復元結果を保存
- 任意で VBA 補助モジュールを使い、再計算や表示値確認を手動補助

## 前提条件

- Windows 上で M365 Excel が利用可能であること
- PowerShell 5.1 以上または PowerShell 7 (`pwsh`) が利用可能であること
- マクロ補助を使う場合は、Excel で VBA 実行が許可されていること

## セキュリティ上の既定動作

- `run_extract.bat` と `run_verify.bat` は、既定で Excel ブックマクロを無効化して開きます
- 信頼済みブックで、どうしても既定動作を変えたい場合だけ `-AllowWorkbookMacros` を明示指定します
- 生成物の絶対パスを減らしたい場合は `run_extract.bat ... -RedactPaths` を使います
- 配布用フォルダ再生成は、既定で `distribution\` 配下のみ安全に削除します

## 主要コマンド

```bat
run_extract.bat "C:\path\to\book.xlsx"
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
run_verify.bat "C:\path\to\book.xlsx"
run_rebuild.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
run_domain_acceptance.bat
run_build_share_package.bat
run_self_test.bat
run_tests.bat
```

`run_extract.bat` は既定で `output/workbook.json`、空またはスキップ状態の `output/styles.json`、`output/manifest.json` を生成します。`run_pack.bat` は既定で `output/llm_package.jsonl` を生成します。`run_rebuild.bat` は既定で `output/rebuilt/*.xlsx` と `output/rebuilt/rebuild_report.json` を生成します。

## 実行例

```bat
run_extract.bat "C:\Data\sample.xlsx"
run_extract.bat "C:\Data\sample.xlsx" -CollectStyles
run_extract.bat "C:\Data\sample.xlsx" -RedactPaths
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json" -ChunkBy range -MaxCells 300
run_verify.bat "C:\Data\sample.xlsx" -WorkbookJsonPath "C:\Work_Codex\Excel2LLM\output\workbook.json"
run_verify.bat "C:\Data\sample.xlsx" -WorkbookJsonPath "C:\Work_Codex\Excel2LLM\output\workbook.json" -AllowWorkbookMacros
run_rebuild.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
run_rebuild.bat "C:\Work_Codex\Excel2LLM\output\workbook.json" -StylesJsonPath "C:\Work_Codex\Excel2LLM\output\styles.json" -OutputPath "C:\Data\rebuilt.xlsx" -Overwrite
run_build_share_package.bat -OutputDir "C:\Temp\Excel2LLM_Share" -AllowOutsideDistribution -ForceCleanOutputDir
```

## 出力ファイル

- `workbook.json`: LLM 投入の正本となるワークブック構造、シート情報、セル情報。`formula` に加えて M365 向けの `formula2`、可能なら `comment_threaded` も保持
- `styles.json`: 低優先の見た目情報。既定ではスキップし、`-CollectStyles` 指定時のみ best effort で生成
- `manifest.json`: 抽出結果の概要、警告、検証状態
- `llm_package.jsonl`: LLM 向けのチャンク化済み JSONL
- `verify_report.json`: 再計算後の差分検証レポート
- `rebuild_report.json`: `workbook.json` からの逆生成結果、警告、復元件数

## 制約

- 初期対応形式は `.xlsx` と `.xlsm` です
- `styles.json` は補助情報のため、既定では取得しません。色や罫線が未取得でも主処理は継続します
- 逆生成の出力形式は常に `.xlsx` です。`has_vba=true` でも VBA 本体は復元しません
- 条件付き書式の見た目そのものは完全再現しません。必要時は VBA 補助または Excel 上での確認を併用してください
