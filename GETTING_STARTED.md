# Excel2LLM はじめに

- 作成日: 2026-03-28 00:20 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-30

## この文書だけ読めば使えます

利用者向けの手順は、この文書に統合しました。
まずはこの文書のとおりに進めてください。

## 迷ったらこれだけで使えます

| やること | どうするか |
| --- | --- |
| 普通に使う | Excel ファイルを `Excel2LLM.bat` にドラッグアンドドロップする |
| 重要な Excel を慎重に扱う | `Excel2LLM.bat` をダブルクリックして「2. Excel を処理して照合もする」を選ぶ |
| うまく動くか確認する | `Excel2LLM.bat` をダブルクリックして「9. 動作確認をする」を選ぶ |

## 使う前に必要なもの

| 項目 | 内容 |
| --- | --- |
| OS | Windows 10 / 11 |
| Excel | Microsoft 365 Excel（デスクトップ版） |
| 追加インストール | 不要 |
| 主な使い方 | `Excel2LLM.bat` へ Excel をドラッグアンドドロップ |

## まずやること

| 手順 | 何をするか | 補足 |
| --- | --- | --- |
| 1 | 配布フォルダをそのまま任意の場所へ置く | フォルダ名は変えてもかまいません |
| 2 | `Excel2LLM.bat` をダブルクリックして「9. 動作確認をする」を選ぶ | 最初の 1 回だけで十分です |
| 3 | `セルフテストが正常終了しました。` と出ることを確認する | これで使える状態です |
| 4 | 自分の Excel を `Excel2LLM.bat` にドラッグアンドドロップする | これが基本の使い方です |

## Excel2LLM.bat のメニューでよく使う番号

| 番号 | 何ができるか | こういうときに使う |
| --- | --- | --- |
| 1 | Excel を処理する | まず普通に使いたい |
| 2 | Excel を処理して照合もする | 重要な資料なので確認もしたい |
| 3 | 見た目情報や追加ルールも含めて抽出する | 色や罫線に加えて、名前定義、入力規則、条件付き書式も見たい |
| 5 | VBA を取り出して LLM 用ファイルも作る | `.xlsm` / `.xlam` のマクロをレビューしたい |
| 6 | 抽出結果と元 Excel を照合する | すでに `workbook.json` があり、あとから確認したい |
| 8 | 抽出結果から Excel を復元する | `workbook.json` から Excel を作り直したい |
| 9 | 最新結果から指示文セットを作る | LLM に貼り付ける文を作りたい |
| 10 | 動作確認をする | 最初の確認や、不具合切り分けをしたい |

## 基本の使い方

### いちばん簡単な使い方

1. 処理したい Excel ファイルを見つける
2. その Excel ファイルを `Excel2LLM.bat` の上へドラッグアンドドロップする
3. 処理が終わるまで待つ
4. 画面に出た出力先フォルダを開く
5. 結果を確認してからキーを押して画面を閉じる

コマンドで実行する場合:

```bat
Excel2LLM.bat "C:\Data\book.xlsx"
```

重要な資料で、Excel との突き合わせもしたい場合:

```bat
Excel2LLM.bat "C:\Data\book.xlsx" -Verify
```

見た目情報も含めて抽出したい場合:

```bat
Excel2LLM.bat -Extract "C:\Data\book.xlsx" -CollectStyles
```

名前定義、入力規則、条件付き書式も含めたい場合:

```bat
Excel2LLM.bat -Extract "C:\Data\book.xlsx" -CollectNamedRanges -CollectDataValidations -CollectConditionalFormats
```

VBA ソースを取り出して LLM に渡したい場合:

```bat
Excel2LLM.bat -MacroExtract "C:\Data\book.xlsm"
```

## 実行すると何が作られるか

実行するたびに、`output` の中へ **新しい実行結果フォルダ** が作られます。
フォルダ名は **ファイル名 + 実行日時** です。

例:

```text
output\estimate_20260328-143500
```

同じ Excel を何回実行しても、前回結果を上書きしにくい作りです。

出力先の見え方の例:

| 例 | 意味 |
| --- | --- |
| `output\estimate_20260328-143500\workbook.json` | 2026-03-28 14:35 に `estimate.xlsx` を処理した結果 |
| `output\estimate_20260328-143500\llm_package.jsonl` | その実行で作られた LLM 用データ |
| `output\latest_run.txt` | 直前に処理した結果フォルダの場所 |

## 出力ファイル一覧

| ファイル / フォルダ | いつできるか | 意味 |
| --- | --- | --- |
| `preflight_report.json` | `Excel2LLM.bat` の通常処理 / `-Extract` / `-Preflight` | 事前チェック結果 |
| `workbook.json` | `Excel2LLM.bat` の通常処理 / `-Extract` | Excel 全体を保存した正本 |
| `styles.json` | `-CollectStyles` を付けたとき | 色や罫線などの補助情報 |
| `workbook.json > named_ranges` | `-CollectNamedRanges` を付けたとき | 名前定義・名前付き範囲 |
| `workbook.json > data_validations` | `-CollectDataValidations` を付けたとき | 入力規則 |
| `workbook.json > conditional_formats` | `-CollectConditionalFormats` を付けたとき | 条件付き書式 |
| `manifest.json` | `Excel2LLM.bat` の通常処理 / `-Extract` | 抽出結果の要約 |
| `llm_package.jsonl` | `Excel2LLM.bat` の通常処理 / `-Pack` | LLM に渡しやすい分割済みデータ |
| `verify_report.json` | `Excel2LLM.bat "..." -Verify` または `Excel2LLM.bat -Verify ...` | Excel との突き合わせ結果 |
| `prompt_bundle\` | `-PromptBundle` | LLM に貼り付ける指示文セット |
| `vba\macro_manifest.json` | `-MacroExtract` | VBA 抽出結果の要約 |
| `vba\modules\` | `-MacroExtract` | `.bas/.cls/.frm` などの可読ソース |
| `vba\vba_llm_package.jsonl` | `-MacroExtract` | VBA を LLM に渡しやすい JSONL |
| `vba\vba_prompt.txt` | `-MacroExtract` | VBA レビュー用の完成文 |
| `rebuilt\` | `Excel2LLM.bat -Rebuild ...` | `workbook.json` から作り直した Excel |

## 最初に見るべきもの

| 見るもの | 何を見るか |
| --- | --- |
| `workbook.json` | Excel の内容が取れているか |
| `llm_package.jsonl` | LLM に渡すデータができているか |
| `verify_report.json` | 差分があるかどうか |
| `preflight_report.json` | 重すぎる・壊れているなどで止まっていないか |

## よく出る言葉

| 用語 | 意味 |
| --- | --- |
| `preflight` | Excel を開く前の事前チェック |
| `workbook.json` | Excel 全体の内容を保存したファイル |
| `JSONL` | 1 行に 1 件ずつデータが入る形式 |
| `チャンク` | LLM に一度に渡すデータのかたまり |
| `prompt bundle` | LLM に貼り付ける指示文セット |
| `VBA 抽出` | `.xlsm/.xlam` からマクロの可読ソースを取り出すこと |
| `verify` | 抽出結果と Excel を突き合わせる確認 |
| `名前定義` | 数式の中で意味のある名前を付けた参照 |
| `入力規則` | セルに入力してよい値のルール |
| `条件付き書式` | 条件に応じて色や見た目を変えるルール |

## 使い分け表

| やりたいこと | 使うもの |
| --- | --- |
| まず普通に使いたい | `Excel2LLM.bat` にドラッグアンドドロップ |
| 重要な Excel を慎重に扱いたい | `Excel2LLM.bat` + `-Verify` |
| 危険な Excel か先に確認したい | `Excel2LLM.bat -Preflight "..."` |
| 見た目情報も含めて抽出したい | `Excel2LLM.bat -Extract "..." -CollectStyles` |
| 名前定義、入力規則、条件付き書式も見たい | `Excel2LLM.bat -Extract "..." -CollectNamedRanges -CollectDataValidations -CollectConditionalFormats` |
| `.xlsm` のマクロをレビューしたい | `Excel2LLM.bat -MacroExtract "..."` |
| LLM に貼り付ける文を作りたい | `Excel2LLM.bat -PromptBundle` |
| `workbook.json` から Excel を作り直したい | `Excel2LLM.bat -Rebuild "...\workbook.json"` |

## Excel2LLM.bat でできること

| やりたいこと | 例 |
| --- | --- |
| 普通に処理する | `Excel2LLM.bat "C:\Data\book.xlsx"` |
| 処理して照合もする | `Excel2LLM.bat "C:\Data\book.xlsx" -Verify` |
| 見た目情報も含めて抽出する | `Excel2LLM.bat -Extract "C:\Data\book.xlsx" -CollectStyles` |
| 名前定義、入力規則、条件付き書式も含めて抽出する | `Excel2LLM.bat -Extract "C:\Data\book.xlsx" -CollectNamedRanges -CollectDataValidations -CollectConditionalFormats` |
| VBA を取り出す | `Excel2LLM.bat -MacroExtract "C:\Data\book.xlsm"` |
| 事前チェックだけ行う | `Excel2LLM.bat -Preflight "C:\Data\book.xlsx"` |
| 抽出結果を照合する | `Excel2LLM.bat -Verify "C:\Data\book.xlsx" -WorkbookJsonPath "output\run\workbook.json"` |
| 抽出結果を分割し直す | `Excel2LLM.bat -Pack "output\run\workbook.json" -ChunkBy range -MaxCells 300` |
| 追加ルールも LLM 用に含める | `Excel2LLM.bat -Pack "output\run\workbook.json" -IncludeNamedRanges -IncludeDataValidations -IncludeConditionalFormats` |
| Excel を復元する | `Excel2LLM.bat -Rebuild "output\run\workbook.json" -StylesJsonPath "output\run\styles.json"` |
| 指示文セットを作る | `Excel2LLM.bat -PromptBundle -Scenario general` |
| 動作確認をする | `Excel2LLM.bat -SelfTest` |

## LLM に渡すまでの流れ

| 手順 | 何をするか |
| --- | --- |
| 1 | `Excel2LLM.bat` で `llm_package.jsonl` を作る |
| 2 | `llm_package.jsonl` をテキストエディタで開く |
| 3 | 必要な行だけをコピーする |
| 4 | ChatGPT などへ貼り付ける |
| 5 | 必要なら `Excel2LLM.bat -PromptBundle` で指示文も作る |

LLM 向けの指示文例は `docs\reference\LLM_PROMPT_FORMATS.md` にあります。

## 困ったとき

| 状況 | まずやること |
| --- | --- |
| うまく動かない | Excel を閉じて、もう一度実行する |
| 途中で止まった | `preflight_report.json` を見る |
| 差分が出た | `verify_report.json` を見る |
| まだだめ | `Excel2LLM.bat -SelfTest` を実行する |

## 最初のうちは覚えなくてよいこと

| 今は気にしなくてよいもの | 理由 |
| --- | --- |
| `tools\advanced\` | 詳細機能用です。通常は `Excel2LLM.bat` だけで足ります |
| `JSONL` の細かい形式 | まずは `llm_package.jsonl` ができれば十分です |
| `styles.json` | 色や罫線が必要なときだけ使います |
| `prompt bundle` の細かい中身 | まずは作成して開ければ十分です |

## 補足

- `Excel2LLM.bat` は、抽出前に自動で事前チェックを行います
- 重すぎる Excel や壊れている疑いがある Excel は、Excel を開く前に停止します
- `Excel2LLM.bat -PromptBundle` は、直前の実行結果フォルダを自動で使います
- `Excel2LLM.bat -MacroExtract` は `.xlsm/.xlam` 専用です
- VBA の可読ソース抽出には、Excel の VBA プロジェクトアクセス許可が必要な場合があります
- `-CollectNamedRanges`, `-CollectDataValidations`, `-CollectConditionalFormats` は、必要なときだけ付けてください
- `Excel2LLM.bat` は処理後に画面を閉じません。結果確認後にキーを押して閉じます
- 詳細機能を直接使いたい場合だけ `tools\advanced\` の `bat` も利用できます
- 絶対パスを減らしたい場合は `-RedactPaths` を使います
