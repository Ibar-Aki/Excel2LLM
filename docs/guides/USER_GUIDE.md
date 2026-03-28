# ユーザーガイド

- 作成日: 2026-03-10 01:30 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## この文書の役割

`Excel2LLM` の詳しい使い方をまとめた手順書です。
最短手順だけ知りたい場合は、先に `MANUAL.md` を読んでください。

## 普段は run_all を使えば十分

普段は次の 1 コマンドで十分です。

```bat
run_all.bat "C:\Data\report.xlsx"
```

重要な資料なら:

```bat
run_all.bat "C:\Data\report.xlsx" -Verify
```

`run_all.bat` は `extract -> pack` をまとめて実行します。
`-Verify` を付けたときだけ、途中で Excel との突き合わせも行います。

## まず理解しておくこと

- `workbook.json`
  - Excel の全シート・全セルの内容を保存したファイル（元データの正確なコピー）
- `preflight_report.json`
  - 抽出前の事前チェック結果
- `llm_package.jsonl`
  - 1 行に 1 つずつチャンクを保存したファイル
- `verify_report.json`
  - Excel との突き合わせ結果
- `styles.json`
  - 色、罫線、配置などの補助情報
- `チャンク`
  - LLM に一度に渡せるサイズに分割したデータのかたまり
- `prompt bundle`
  - LLM に貼り付けるための指示文セット

## 基本の流れ

まずは `run_all.bat` を使えば十分です。
必要に応じて、以下の個別コマンドを使い分けます。

### 0. 事前チェックで危険ファイルを止める

`extract` と `run_all` は、抽出前に必ず `preflight`（事前チェック）を実行します。

```bat
run_preflight.bat "C:\Data\report.xlsx"
```

停止条件の例:

- ファイルサイズが 200MB を超える
- 推定総セル数が 500 万を超える
- 単一シートの推定セル数が 200 万を超える
- `xlsx/xlsm` の ZIP 構造や必須 XML が壊れている

### 1. Excel を JSON に変換する

```bat
run_extract.bat "C:\Data\report.xlsx"
```

主な出力:

- `output\preflight_report.json`
- `output\workbook.json`
- `output\styles.json`
- `output\manifest.json`

### 2. LLM に渡すチャンクを作る

```bat
run_pack.bat "output\workbook.json"
```

主な出力:

- `output\llm_package.jsonl`

### 3. 重要な資料なら検証する

```bat
run_verify.bat "C:\Data\report.xlsx" -WorkbookJsonPath "output\workbook.json"
```

主な出力:

- `output\verify_report.json`

### 4. 必要なら Excel を作り直す

```bat
run_rebuild.bat "output\workbook.json"
```

### 5. prompt bundle を作る

```bat
run_prompt_bundle.bat -Scenario general
```

主な出力:

- `output\prompt_bundle\prompt_*.txt`
- `output\prompt_bundle\prompt_bundle_manifest.json`

## 各コマンドの使い分け

### extract

用途:

- Excel の内容を壊れにくい JSON にしたい
- 数式、表示値、結合セルを保持したい

例:

```bat
run_extract.bat "C:\Data\sample.xlsx"
run_extract.bat "C:\Data\sample.xlsx" -CollectStyles
run_extract.bat "C:\Data\sample.xlsx" -RedactPaths
run_extract.bat "C:\Data\sample.xlsx" -Sheets Summary,Calc
run_extract.bat "C:\Data\sample.xlsx" -Sheets Summary,Calc -ExcludeSheets Calc
```

補足:

- `run_preflight.bat`
  - Excel を開かずに危険ファイルを止めたいときに使う
- `-Sheets`
  - 指定したシートだけ抽出
- `-ExcludeSheets`
  - 指定したシートを除外
- `-RedactPaths`
  - 出力に絶対パスを残しにくくする
- `-CollectStyles`
  - `styles.json` を取得する

### pack

用途:

- 大きい Excel を LLM 向けに小分けしたい

例:

```bat
run_pack.bat "output\workbook.json"
run_pack.bat "output\workbook.json" -ChunkBy range -MaxCells 300
run_pack.bat "output\workbook.json" -ChunkBy sheet -IncludeStyles
```

補足:

- `-ChunkBy sheet`
  - 行のまとまりを保ちやすい
- `-ChunkBy range`
  - セル数優先で細かく分ける

### verify

用途:

- 抽出した JSON と Excel の再計算結果を突き合わせたい
- `run_all -Verify` では、この verify が途中で自動実行される

例:

```bat
run_verify.bat "C:\Data\sample.xlsx" -WorkbookJsonPath "output\workbook.json"
run_verify.bat "C:\Data\sample.xlsx" -WorkbookJsonPath "output\workbook.json" -RedactPaths
```

### rebuild

用途:

- `workbook.json` から `.xlsx` を作り直したい

例:

```bat
run_rebuild.bat "output\workbook.json"
run_rebuild.bat "output\workbook.json" -StylesJsonPath "output\styles.json" -OutputPath "C:\Data\rebuilt.xlsx" -Overwrite
```

### prompt bundle

用途:

- LLM にそのまま貼りやすい prompt テキストを作りたい
- `prompt bundle` は「LLM に貼り付けるための指示文セット」の意味

例:

```bat
run_prompt_bundle.bat -Scenario general
run_prompt_bundle.bat -Scenario mechanical
run_prompt_bundle.bat -Scenario accounting
```

## ヘルプの見方

各 `run_*.bat` は、引数なしでも使い方を表示します。

```bat
run_all.bat -h
run_extract.bat -h
run_preflight.bat -h
run_pack.bat --help
run_prompt_bundle.bat -h
run_verify.bat /?
run_rebuild.bat -h
```

## 実行後に見るべきこと

### extract のあと

- コンソールの `Excel2LLM 抽出結果`
- `preflight_report.json` の `warnings`
- `manifest.json` の `warnings`
- `workbook.json` の `sheet_count`, `formula_count`
- 表示された `次のおすすめ`

### pack のあと

- コンソールの `パッキング結果`
- `llm_package.jsonl` のチャンク数
- 表示された `次のおすすめ`

### verify のあと

- コンソールの `検証結果`
- `verify_report.json` の `mismatch_count`
- 差分がある場合は `verify_report.json` を開く

## preflight で止まったとき

画面に停止理由が出たら、まず次を確認します。

1. 対象シートを絞る
2. 不要列を削る
3. 集計用シートだけを別ブック化する

`preflight_report.json` を見ると、停止理由、警告、推定セル数、最大シートが分かります。

## エラー時の見方

主要コマンドは、失敗時に次の番号付き案内を表示します。

1. Excel を閉じる
2. コマンドをもう一度実行する
3. まだダメなら `run_self_test.bat`

## LLM へ渡すときの考え方

- まず `llm_package.jsonl` から対象チャンクを選ぶ
- 数式確認なら `formula` / `formula2` を重視する
- 表示上の見え方が重要なら `text` を参照する
- 値の比較や統計は `value2` を基準にする

用途別の指示文は `..\reference\LLM_PROMPT_FORMATS.md` を使ってください。

## よくあるつまずき

### 何から始めればいいかわからない

1. `run_all.bat "対象.xlsx"`
2. `output\llm_package.jsonl` を使う

### コマンドが失敗した

1. まず `run_*.bat -h` で使い方を確認
2. Excel を閉じて再実行
3. `output\manifest.json` または `output\verify_report.json` を確認
4. まだ不明なら `run_self_test.bat` を実行

### 配布用にまとめたい

`SHARE_PACKAGE.md` を参照してください。
