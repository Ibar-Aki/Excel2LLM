# Excel2LLM クイックスタート

- 作成日: 2026-03-11 02:35 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## この文書の役割

この文書は、初回 3 分で使い始めるための短い手順だけをまとめたものです。
詳しい説明は `USER_GUIDE.md` に集約しています。

## 最短手順

### 1. まとめて実行する

```bat
run_all.bat "C:\Data\book.xlsx"
```

このとき、抽出前に自動で `preflight` が走ります。重すぎる Excel や破損疑いのある Excel はここで停止します。

### 2. 重要なら検証も入れる

```bat
run_all.bat "C:\Data\book.xlsx" -Verify
```

### 3. prompt bundle が欲しいとき

```bat
run_prompt_bundle.bat -Scenario general
```

## まず見るファイル

- `output\workbook.json`
  - Excel の内容を保存した正本
- `output\preflight_report.json`
  - 事前チェックの結果
- `output\llm_package.jsonl`
  - LLM に渡しやすい分割済みデータ
- `output\verify_report.json`
  - Excel との突き合わせ結果
- `output\prompt_bundle\prompt_*.txt`
  - LLM に貼り付けるための prompt テキスト

## 困ったとき

- 使い方を見たい
  - `run_all.bat -h`
  - `run_extract.bat -h`
  - `run_preflight.bat -h`
  - `run_pack.bat -h`
  - `run_prompt_bundle.bat -h`
  - `run_verify.bat -h`
  - `run_rebuild.bat -h`
- 詳しい手順を見たい
  - `USER_GUIDE.md`
- LLM への指示文を見たい
  - `..\reference\LLM_PROMPT_FORMATS.md`

## 補足

- `extract` と `verify` は、既定でマクロを無効化して Excel を開きます
- `extract` の前には必ず preflight が走り、危険なブックは停止します
- 絶対パスを減らしたい場合は `-RedactPaths` を使います
- 特定シートだけ抽出したい場合は `-Sheets` と `-ExcludeSheets` を使います
