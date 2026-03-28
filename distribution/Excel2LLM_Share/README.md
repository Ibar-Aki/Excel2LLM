# Excel2LLM

- 作成日: 2026-03-10 00:55 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

Windows と M365 Excel を前提に、Excel ブックを LLM 向けの JSON / JSONL に変換するツールです。

## 3ステップで使う

### 1. まとめて実行する

```bat
run_all.bat "C:\Data\book.xlsx"
```

### 2. 重要なら検証も入れる

```bat
run_all.bat "C:\Data\book.xlsx" -Verify
```

### 3. 必要なら prompt bundle を作る

```bat
run_prompt_bundle.bat -Scenario general
```

主に使う出力:

- `output\preflight_report.json`
- `output\workbook.json`
- `output\llm_package.jsonl`
- `output\verify_report.json`

## 最初に読む文書

- `GETTING_STARTED.md`
  - 配布先の人が最初に読む 1 ページ
- `docs/guides/MANUAL.md`
  - 初回 3 分で使うための短いクイックスタート
- `docs/guides/USER_GUIDE.md`
  - 詳しい手順、オプション、出力の見方、トラブル対応
- `docs/reference/LLM_PROMPT_FORMATS.md`
  - LLM へ渡すときの用途別テンプレート
- `docs/guides/SHARE_PACKAGE.md`
  - 他の人へ渡すための配布用フォルダの作り方
- `docs/README.md`
  - 文書全体の案内

## 主要コマンド

```bat
run_all.bat "C:\path\to\book.xlsx"
run_extract.bat "C:\path\to\book.xlsx"
run_preflight.bat "C:\path\to\book.xlsx"
run_pack.bat "output\workbook.json"
run_prompt_bundle.bat -Scenario general
run_verify.bat "C:\path\to\book.xlsx" -WorkbookJsonPath "output\workbook.json"
run_rebuild.bat "output\workbook.json"
```

各 `run_*.bat` は、引数なし、`-h`、`--help`、`/?` のときに使い方を表示します。

## できること

- 複数シートのセル値、表示値、数式、結合セルを `workbook.json` に保存
- 色や罫線などの補助情報を `styles.json` に分離保存
- LLM 向けの `llm_package.jsonl` を生成
- `prompt_*.txt` の prompt bundle を生成
- 抽出結果と Excel 再計算結果の差分を `verify_report.json` で確認
- `workbook.json` から `.xlsx` を逆生成

## セキュリティ上の既定動作

- `extract` の前に必須の preflight を行い、重すぎる Excel や破損疑いのある Excel は抽出開始前に停止します
- `extract` と `verify` は、既定で Excel ブックマクロを無効化して開きます
- 絶対パスを減らしたい場合は `-RedactPaths` を使います
- 配布用フォルダ再生成は、既定で `distribution\` 配下のみ安全に削除します

## 補足

- 対応形式は `.xlsx` と `.xlsm`
- 色や罫線は補助情報であり、主処理の成功条件ではありません
- 逆生成の出力は常に `.xlsx` です
