# ユーザーガイド

- 作成日: 2026-03-10 01:30 JST
- 作成者: Codex (GPT-5)

## これは何か

`Excel2LLM` は、Excel ブックをそのまま LLM に投げる代わりに、まず欠落の少ない JSON へ変換し、その後に LLM 向けのチャンクへ整形するためのツールです。

このツールが特に向いているケース:

- 数式を落とさずに LLM へ渡したい
- 複数シートをまとめて扱いたい
- 結合セルや表示値も保持したい
- 1 回のプロンプトに入りきらない Excel を安全に分割したい

## 最初に理解しておくこと

- 主データは `workbook.json` です
- LLM に渡す実用データは `llm_package.jsonl` です
- 色や罫線は低優先で、必要なときだけ `styles.json` を使います
- `verify_report.json` は「抽出した JSON と、Excel 再計算後の実データに差がないか」を見るための確認用です

## 最短手順

### 1. Excel を JSON に変換する

```bat
run_extract.bat "C:\Data\report.xlsx"
```

生成される主なファイル:

- `output\workbook.json`
- `output\styles.json`
- `output\manifest.json`

### 2. LLM に渡すチャンクを作る

```bat
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

生成されるファイル:

- `output\llm_package.jsonl`

### 3. 必要なら Excel と突き合わせる

```bat
run_verify.bat "C:\Data\report.xlsx"
```

生成されるファイル:

- `output\verify_report.json`

## どのコマンドを使うべきか

### まず抽出だけしたい

```bat
run_extract.bat "C:\Data\book.xlsx"
```

用途:

- 数式や結合セルを欠落なく保存したい
- まだ LLM に渡す単位を決めていない
- 後工程で別の変換をしたい

### LLM に渡しやすい単位に分けたい

```bat
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json" -ChunkBy range -MaxCells 300
```

用途:

- 大きな表を小さく分けたい
- RAG やエージェント処理の入力にしたい
- シート全体では大きすぎる

### 抽出結果が信用できるか確認したい

```bat
run_verify.bat "C:\Data\book.xlsx" -WorkbookJsonPath "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

用途:

- 数式の再計算結果が心配
- Excel 側の表示値とのズレを確認したい
- 重要な資料を LLM に渡す前に検証したい

## よく使うオプション

### `-ChunkBy sheet`

シート単位で考えながら、必要に応じて行境界を保って分割します。

向いているケース:

- 1 シートが 1 トピックになっている
- 表の行まとまりを壊したくない

### `-ChunkBy range`

セル数ベースで機械的に分割します。

向いているケース:

- とにかく入力サイズを一定にしたい
- 多少レンジがまたがってもよい

### `-MaxCells`

1 チャンクに入れる最大セル数です。

目安:

- 小さめにしたい: `100` から `300`
- やや大きめでもよい: `300` から `800`

### `-CollectStyles`

色や罫線などの見た目情報を `styles.json` に出したいときだけ使います。

```bat
run_extract.bat "C:\Data\book.xlsx" -CollectStyles
```

注意:

- 既定では style を取りません
- style 取得は遅くなりやすいです
- まずは style なしで運用し、必要な案件だけ付けるのが安全です

## 出力ファイルの見方

### `workbook.json`

最重要ファイルです。

主に見る項目:

- `workbook.sheet_count`
- `sheets[].sheet_name`
- `cells[].address`
- `cells[].value2`
- `cells[].text`
- `cells[].formula`
- `cells[].formula2`
- `cells[].merge_area`

使いどころ:

- 監査用の正本
- 後続変換の元データ
- 「どのセルに何が入っていたか」の確認

### `llm_package.jsonl`

LLM に流し込みやすい実用ファイルです。1 行が 1 チャンクです。

主に見る項目:

- `chunk_id`
- `sheet_name`
- `range`
- `payload`
- `formula_cells`

使いどころ:

- RAG の投入
- エージェントの逐次処理
- バッチ要約

### `manifest.json`

処理の概要が入っています。

主に見る項目:

- `status`
- `warnings`
- `formula_count`
- `merged_range_count`
- `style_export_status`
- `verify_status`

### `verify_report.json`

Excel 再計算後との差分です。

主に見る項目:

- `status`
- `mismatch_count`
- `mismatches[]`

## LLM への渡し方

### 方法 1: まず人が確認してから渡す

1. `workbook.json` を作る
2. `llm_package.jsonl` を作る
3. 対象チャンクを選ぶ
4. そのチャンクだけ LLM に渡す

向いているケース:

- 重要資料
- 財務、契約、見積もり
- 数式の意味まで確認したい案件

### 方法 2: 全チャンクを順番に処理させる

1. `llm_package.jsonl` を作る
2. 1 行ずつ LLM やエージェントに渡す
3. 最後に統合要約を作る

向いているケース:

- 大きいブック
- 多数シート
- 定期処理

### 方法 3: `workbook.json` を検索用データにする

1. `workbook.json` を保存する
2. シート名、セル番地、式、表示値で検索できる形に載せる
3. 必要箇所だけ LLM に渡す

向いているケース:

- 監査用途
- FAQ ボット
- 社内検索

## 推奨プロンプト

### シート構造を理解させたい

```text
以下は Excel の正規化データです。sheet_name、range、cells を見て、各シートの役割と主要な数式を説明してください。数式セルは formula または formula2 を優先して参照してください。
```

### 数式の意味を説明させたい

```text
以下の Excel チャンクに含まれる formula / formula2 を読み、各数式が何を計算しているか、入力セルと出力セルの関係を日本語で整理してください。
```

### 表を要約させたい

```text
以下の Excel チャンクを読み、列の意味、重要な行、異常値の可能性を要約してください。text と value2 の違いがある場合はその差も指摘してください。
```

## トラブルシューティング

### `styles.json` が空

正常です。既定では style を取りません。

必要なら:

```bat
run_extract.bat "C:\Data\book.xlsx" -CollectStyles
```

### `verify_report.json` で差分が出た

確認ポイント:

- Excel を開いた直後に再計算が必要なブックか
- 外部参照や volatile 関数があるか
- 手入力後に保存していない変更があったか

### 出力が大きすぎる

対策:

- `run_pack.bat` で `-MaxCells` を小さくする
- `-ChunkBy range` を使う
- 必要シートだけ別ファイルで処理する

### 期待したセルが見つからない

確認ポイント:

- そのセルが `UsedRange` の外ではないか
- 非表示でも実際に使用されているセルか
- 数式結果だけでなく `formula` や `formula2` も見ているか

## 運用のおすすめ

- まずは `run_extract.bat` と `run_pack.bat` だけで始める
- 重要案件だけ `run_verify.bat` を併用する
- 色や罫線は必要になってから `-CollectStyles` を使う
- LLM には `workbook.json` 全量ではなく、`llm_package.jsonl` の必要チャンクだけを渡す
