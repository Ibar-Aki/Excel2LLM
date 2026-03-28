# Excel2LLM はじめに

- 作成日: 2026-03-28 00:20 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## 最初にやること

1. このフォルダをそのまま任意の場所へ置く
2. 動作確認をする

```bat
run_self_test.bat
```

3. 自分の Excel で試す

```bat
run_all.bat "C:\Data\book.xlsx"
```

`run_all` は抽出前に自動で `preflight` を行います。重すぎる Excel や破損疑いのある Excel はここで止まります。

## 何ができるか

- Excel を `workbook.json` に変換する
- LLM に渡しやすい `llm_package.jsonl` を作る
- 必要なら `verify_report.json` で Excel と突き合わせる

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

## 困ったとき

1. Excel を閉じる
2. もう一度コマンドを実行する
3. まだダメなら `run_self_test.bat`

詳しい手順:

- `docs\guides\MANUAL.md`
- `docs\guides\USER_GUIDE.md`
