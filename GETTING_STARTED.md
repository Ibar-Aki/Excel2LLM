# Excel2LLM はじめに

- 作成日: 2026-03-28 00:20 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## 動かすために必要なもの

- Windows 10 / 11
- Microsoft 365 Excel（デスクトップ版）
- PowerShell（Windows に標準搭載）

## この文書の役割

この文書は、配布フォルダを受け取った人が最初の 1 回だけ読むための案内です。
毎回の実行で参照する短い早見表は `docs\guides\MANUAL.md` にまとめています。

## 最初にやること

1. このフォルダをそのまま任意の場所へ置く
2. 動作確認をする

```bat
run_self_test.bat
```

`セルフテストが正常終了しました。` と表示されたら成功です。

3. 自分の Excel で試す

```bat
run_all.bat "C:\Data\book.xlsx"
```

`run_all` は抽出前に自動で `preflight` を行います。重すぎる Excel や破損疑いのある Excel はここで止まります。

実行後、`output` フォルダに次ができていれば成功です。

- `workbook.json`
  - Excel 全体の内容を保存したファイル
- `llm_package.jsonl`
  - LLM に渡せるサイズへ分割したファイル

💡 ヒント: Excel ファイルを `run_all.bat` にドラッグ＆ドロップしても実行できます。

## 何ができるか

- Excel を `workbook.json` に変換する
- LLM に渡しやすい `llm_package.jsonl` を作る
- 必要なら `verify_report.json` で Excel と突き合わせる

## よく出る言葉

- `preflight`
  - Excel を開く前に行う事前チェックです
- `workbook.json`
  - Excel の全シート・全セルの内容を保存したファイルです
- `JSONL`
  - 1 行に 1 つずつデータを入れた軽量形式です
- `チャンク`
  - LLM に一度に渡せるサイズに分割したデータのかたまりです
- `prompt bundle`
  - LLM に貼り付けるための指示文セットです

## まず見るファイル

- `output\workbook.json`
- `output\preflight_report.json`
- `output\llm_package.jsonl`
- `output\verify_report.json`

## 便利なコマンド

```bat
run_all.bat "C:\Data\book.xlsx"
run_all.bat "C:\Data\book.xlsx" -Verify
run_preflight.bat "C:\Data\book.xlsx"
run_prompt_bundle.bat -Scenario general
```

## LLM に渡すには

1. `output\llm_package.jsonl` をテキストエディタで開く
2. 必要な行だけを選んでコピーする
3. ChatGPT などに貼り付けて質問する

詳しいテンプレートは `docs\reference\LLM_PROMPT_FORMATS.md` にあります。

## 困ったとき

1. Excel を閉じる
2. もう一度コマンドを実行する
3. まだダメなら `run_self_test.bat`

詳しい手順:

- `docs\guides\MANUAL.md`
- `docs\guides\USER_GUIDE.md`
