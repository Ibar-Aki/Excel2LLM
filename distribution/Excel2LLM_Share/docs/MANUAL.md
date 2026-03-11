# Excel2LLM ユーザーマニュアル

- 作成日: 2026-03-11 02:35 JST
- 作成者: Codex (GPT-5)

## このマニュアルの目的

この文書は、`Excel2LLM` を初めて使う人が、迷わず実行できるように作った手順書です。

「何をするツールなのか」「どのコマンドを打てばよいのか」「どのファイルを見ればよいのか」を、できるだけ順番に説明します。

## まず結論

通常は、次の 3 ステップだけ覚えれば使えます。

### 1. Excel を読み出す

```bat
run_extract.bat "C:\Data\book.xlsx"
```

### 2. LLM に渡しやすい形に分割する

```bat
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

### 3. 必要なら Excel と突き合わせる

```bat
run_verify.bat "C:\Data\book.xlsx"
```

これで、次のファイルが主に使われます。

- `output\workbook.json`
- `output\llm_package.jsonl`
- `output\verify_report.json`

## このツールで何ができるか

このツールは、Excel をそのまま LLM に渡す代わりに、いったん壊れにくい JSON に変換します。

特に強いのは次の点です。

- 複数シートをまとめて扱える
- 数式を `formula` / `formula2` として保持できる
- セルの表示値 `text` と内部値 `value2` を分けて保持できる
- 結合セル情報を保持できる
- 大きい表を LLM 向けに小分けできる
- 必要なら JSON から `.xlsx` を作り直せる

## 最初に理解しておくべきファイル

### `workbook.json`

最重要ファイルです。Excel の内容を、できるだけ欠落少なく保存した正本です。

迷ったら、まずこれを残してください。

### `llm_package.jsonl`

LLM に渡すための実用ファイルです。

`workbook.json` をそのまま LLM に入れると大きすぎることが多いため、このファイルを使って分割投入します。

### `verify_report.json`

抽出した JSON と、Excel を再計算した結果に差がないかを見る確認用ファイルです。

重要資料では、この確認を入れるのが安全です。

### `styles.json`

色、罫線、配置などの補助情報です。

これは低優先です。値や数式のほうが重要です。

### `rebuild_report.json`

`workbook.json` から Excel を逆生成したときの結果です。

復元時の警告や、VBA が戻っていないことなどをここで確認します。

## どんな順番で使うのか

### パターン 1: 普通に LLM へ渡したい

この使い方が基本です。

1. `run_extract.bat` で `workbook.json` を作る
2. `run_pack.bat` で `llm_package.jsonl` を作る
3. 必要なチャンクだけ LLM に渡す

### パターン 2: 重要な Excel を安全に扱いたい

数値や表示値が重要なら、確認を 1 ステップ増やします。

1. `run_extract.bat`
2. `run_verify.bat`
3. 差分がなければ `run_pack.bat`
4. LLM に渡す

### パターン 3: JSON から Excel を戻したい

1. `run_rebuild.bat`
2. 必要なら `rebuild_report.json` を確認
3. 戻した Excel を目視確認

## いちばん簡単な使い方

### 手順 1: Excel を JSON に変換する

```bat
run_extract.bat "C:\Data\report.xlsx"
```

これで主に次が作られます。

- `output\workbook.json`
- `output\styles.json`
- `output\manifest.json`

ここで大事なのは `workbook.json` です。

### 手順 2: LLM に渡せる形にする

```bat
run_pack.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

これで次が作られます。

- `output\llm_package.jsonl`

このファイルは 1 行が 1 チャンクです。

大きい Excel の場合は、このチャンクを 1 個ずつ LLM に渡します。

### 手順 3: 重要なファイルなら検証する

```bat
run_verify.bat "C:\Data\report.xlsx"
```

これで次が作られます。

- `output\verify_report.json`

`mismatch_count` が 0 なら、抽出結果と Excel 再計算結果に差がなかった、という意味です。

### 手順 4: 必要なら Excel を作り直す

```bat
run_rebuild.bat "C:\Work_Codex\Excel2LLM\output\workbook.json"
```

これで次が作られます。

- `output\rebuilt\*.xlsx`
- `output\rebuilt\rebuild_report.json`

## コマンドの使い分け

### `run_extract.bat`

役割:

- Excel を読む
- シート、セル、数式、結合セル、表示値を JSON にする

使う場面:

- まず正本を作りたい
- LLM に渡す前の元データを残したい
- 後で監査や比較をしたい

### `run_pack.bat`

役割:

- `workbook.json` を LLM 向けに分割する

使う場面:

- Excel が大きい
- 1 回のプロンプトに入りきらない
- シート単位または範囲単位で分けたい

### `run_verify.bat`

役割:

- 抽出した JSON と Excel の再計算結果を比較する

使う場面:

- 財務、見積もり、契約、集計表などの重要ファイル
- 数式の再計算が気になる
- Excel を開くと値が変わることがある

### `run_rebuild.bat`

役割:

- `workbook.json` から `.xlsx` を逆生成する

使う場面:

- JSON 化したデータを Excel に戻したい
- round-trip の確認をしたい
- LLM 処理後の内容を再配布したい

## よく使うオプション

### `-CollectStyles`

色や罫線を `styles.json` に出したいときに使います。

```bat
run_extract.bat "C:\Data\book.xlsx" -CollectStyles
```

このオプションは、通常は不要です。

まずは style なしで始めて、必要な案件だけ使うのが安全です。

### `-ChunkBy sheet`

シートのまとまりを重視して分割します。

向いているケース:

- 1 シートが 1 テーマ
- 行の文脈を壊したくない

### `-ChunkBy range`

セル数ベースで細かく分割します。

向いているケース:

- とにかくサイズを小さくしたい
- 50 行 100 列のような大きい表を扱う

### `-MaxCells`

1 チャンクあたりの最大セル数です。

目安:

- かなり小さくしたいなら `100` から `200`
- 普通は `200` から `400`
- 多少大きくてもよいなら `400` から `800`

### `-Overwrite`

逆生成先の `.xlsx` を上書きしたいときに使います。

```bat
run_rebuild.bat "C:\Work_Codex\Excel2LLM\output\workbook.json" -Overwrite
```

## 何を LLM に渡せばよいか

### 原則

原則として、LLM には `llm_package.jsonl` の必要チャンクだけを渡してください。

`workbook.json` は正本であり、保管や監査には向きますが、毎回そのまま全文を渡すには大きすぎることがあります。

### おすすめの流れ

1. `run_pack.bat` を実行する
2. `llm_package.jsonl` から対象チャンクを選ぶ
3. そのチャンクだけ LLM に貼る
4. 必要なら複数チャンクの結果を最後に統合する

## LLM への質問例

### シートの役割を説明してほしい

```text
以下は Excel のチャンクです。sheet_name、range、cells を見て、このシートの目的、主要な列、重要な数式を日本語で説明してください。
```

### 数式の意味を説明してほしい

```text
以下の Excel チャンクに含まれる formula と formula2 を読み、各数式が何を計算しているか、入力と出力の関係が分かるように整理してください。
```

### おかしな値を探してほしい

```text
以下の Excel チャンクを確認し、空欄、不自然な値、重複、列の意味に対して違和感のあるデータを指摘してください。text と value2 が違う場合は、その理由も推定してください。
```

## 出力ファイルの見方

### `manifest.json`

処理全体の要約です。

よく見る項目:

- `status`
- `warnings`
- `formula_count`
- `merged_range_count`
- `style_export_status`
- `verify_status`

### `workbook.json`

中身を深く確認したいときに見ます。

よく見る項目:

- `sheets[].sheet_name`
- `cells[].address`
- `cells[].value2`
- `cells[].text`
- `cells[].formula`
- `cells[].formula2`
- `cells[].merge_area`

### `llm_package.jsonl`

LLM に渡す単位を確認したいときに見ます。

よく見る項目:

- `chunk_id`
- `sheet_name`
- `range`
- `payload`
- `formula_cells`

### `verify_report.json`

差分の有無を見るときに使います。

よく見る項目:

- `status`
- `mismatch_count`
- `mismatches`

### `rebuild_report.json`

逆生成の結果確認に使います。

よく見る項目:

- `status`
- `warnings`
- `output_path`
- `restored_sheets`
- `restored_cells`
- `restored_formulas`
- `restored_styles`

## よくある失敗と対処

### `styles.json` がほぼ空

異常ではありません。

既定では style は重要視していないため、十分に取られないことがあります。

必要なら次を使います。

```bat
run_extract.bat "C:\Data\book.xlsx" -CollectStyles
```

### 出力が大きすぎる

対処:

- `-ChunkBy range` を使う
- `-MaxCells` を小さくする
- 必要なシートだけ別々に処理する

### 数値はあるのに見た目が違う

このツールでは、値と数式が最優先です。

色、罫線、条件付き書式の見た目は補助扱いです。見た目まで必要な案件だけ `styles.json` を使ってください。

### `verify_report.json` に差分が出る

確認ポイント:

- 開き直しで再計算されるブックではないか
- 外部参照があるか
- volatile 関数があるか
- 手元の Excel が未保存変更を含んでいないか

### 逆生成したら `.xlsm` にならない

仕様です。

逆生成の出力は常に `.xlsx` です。VBA 本体は戻しません。

## 制約

- 対応形式は主に `.xlsx` と `.xlsm`
- VBA 本体は復元しない
- 色や罫線は低優先で best effort
- threaded comment は完全再現ではなく通常コメントへ寄せることがある
- `UsedRange` 外の情報は取得対象外になることがある

## 困ったときに見る順番

迷ったら、次の順で確認してください。

1. `README.md`
2. この `MANUAL.md`
3. `docs/USER_GUIDE.md`
4. `docs/USE_CASES.md`
5. `docs/FORMAT.md`

## 最後に

このツールの基本は単純です。

- まず `run_extract.bat`
- 次に `run_pack.bat`
- 重要なら `run_verify.bat`
- 戻したいときだけ `run_rebuild.bat`

この順番を守れば、大きく迷わず使えます。
