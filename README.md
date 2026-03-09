# Excel2LLM

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-10

Windows と M365 Excel を前提に、Excel ブックを LLM 向けの正規化 JSON と JSONL に変換するツール群です。追加インストールは前提にせず、`PowerShell`、`Excel COM`、`VBA`、`bat` だけで動作します。

## ディレクトリ構成

- `scripts/`: 抽出、パック、検証、サンプル生成の PowerShell スクリプト
- `output/`: 生成される `workbook.json`、`styles.json`、`llm_package.jsonl` などの出力先
- `docs/`: 運用メモ、データ形式、VBA 補助の説明
- `templates/`: 任意で Excel に取り込める VBA テンプレート

## できること

- 複数シートのセル内容、数式、表示値、結合セル情報を `workbook.json` に正規化保存
- 色、罫線、配置などの見た目情報を `styles.json` に分離保存
- LLM にそのまま流し込みやすい `llm_package.jsonl` を生成
- Excel 再計算後の差分確認を `verify_report.json` と `manifest.json` に保存
- 任意で VBA 補助モジュールを使い、再計算や表示値確認を手動補助

## 前提条件

- Windows 上で M365 Excel が利用可能であること
- PowerShell 5.1 以上または PowerShell 7 (`pwsh`) が利用可能であること
- マクロ補助を使う場合は、Excel で VBA 実行が許可されていること

## 主要コマンド

```bat
run_extract.bat "C:\path\to\book.xlsx"
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
run_verify.bat "C:\path\to\book.xlsx"
run_self_test.bat
```

`run_extract.bat` は既定で `output/workbook.json`、空またはスキップ状態の `output/styles.json`、`output/manifest.json` を生成します。`run_pack.bat` は既定で `output/llm_package.jsonl` を生成します。

## 実行例

```bat
run_extract.bat "C:\Data\sample.xlsx"
run_extract.bat "C:\Data\sample.xlsx" -CollectStyles
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json" -ChunkBy range -MaxCells 300
run_verify.bat "C:\Data\sample.xlsx" -WorkbookJsonPath "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

## 出力ファイル

- `workbook.json`: LLM 投入の正本となるワークブック構造、シート情報、セル情報
- `workbook.json`: LLM 投入の正本となるワークブック構造、シート情報、セル情報。`formula` に加えて M365 向けの `formula2`、可能なら `comment_threaded` も保持
- `styles.json`: 低優先の見た目情報。既定ではスキップし、`-CollectStyles` 指定時のみ best effort で生成
- `manifest.json`: 抽出結果の概要、警告、検証状態
- `llm_package.jsonl`: LLM 向けのチャンク化済み JSONL
- `verify_report.json`: 再計算後の差分検証レポート

## 制約

- 初期対応形式は `.xlsx` と `.xlsm` です
- `styles.json` は補助情報のため、既定では取得しません。色や罫線が未取得でも主処理は継続します
- 条件付き書式の見た目そのものは完全再現しません。必要時は VBA 補助または Excel 上での確認を併用してください
