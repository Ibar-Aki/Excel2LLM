# 抽出できる情報 / できない情報

- 作成日: 2026-03-29 18:30 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-29

## この文書の目的

この文書は、Excel2LLM が元の Excel から何を取得するか、何を取得しないかを分かりやすく整理した一覧です。

特に次の判断に使えます。

- このツールでそのまま分析できるか
- 追加オプションを付けるべきか
- 取得できない情報があるので別の確認が必要か

## 既定で取得する情報

| 分類 | 取得有無 | 出力先 | 内容 | 復元対応 |
| --- | --- | --- | --- | --- |
| シート情報 | 取得する | `workbook.json` | シート名、順序、表示状態、使用範囲 | 一部対応 |
| セル値 | 取得する | `workbook.json` | `value2` と `text` | 対応 |
| 数式 | 取得する | `workbook.json` | `formula` と `formula2` | 対応 |
| 表示書式 | 取得する | `workbook.json` | `number_format` | 対応 |
| 結合セル | 取得する | `workbook.json` | `merge_area`, `merged_ranges` | 対応 |
| コメント | 取得する | `workbook.json` | 通常コメント、threaded comment の本文 | 一部対応 |
| ハイパーリンク | 取得する | `workbook.json` | リンク先、表示文字 | 一部対応 |
| hidden 行列 | 取得する | `workbook.json` | 非表示行、非表示列 | 対応 |
| freeze panes | 取得する | `workbook.json` | 枠固定の状態 | 対応 |

## オプションで取得する情報

| 分類 | 取得オプション | 既定 | 出力先 | 内容 | 復元対応 |
| --- | --- | --- | --- | --- | --- |
| 色・罫線などの補助書式 | `-CollectStyles` | 取得しない | `styles.json` | 塗りつぶし色、文字色、整列、折り返し、罫線 | 一部対応 |
| 名前定義・名前付き範囲 | `-CollectNamedRanges` | 取得しない | `workbook.json` | 定義名、スコープ、参照式、単純参照先 | 未対応 |
| 入力規則 | `-CollectDataValidations` | 取得しない | `workbook.json` | 種別、対象範囲、候補式、入力 / エラー文言 | 未対応 |
| 条件付き書式 | `-CollectConditionalFormats` | 取得しない | `workbook.json` | 対象範囲、ルール種別、条件式、優先度 | 未対応 |

## LLM 用ファイルに含められる情報

| 分類 | パック時オプション | `llm_package.jsonl` に入る場所 | 説明 |
| --- | --- | --- | --- |
| セル本体 | 既定で含む | `payload.cells[]` | 値、表示値、数式、コメントなど |
| 補助書式 | `-IncludeStyles` | `payload.cells[].style` | `styles.json` からセル単位で付加 |
| 名前定義・名前付き範囲 | `-IncludeNamedRanges` | `payload.metadata.named_ranges` | チャンクに関係する定義名を付加 |
| 入力規則 | `-IncludeDataValidations` | `payload.metadata.data_validations` | チャンク範囲に重なる入力規則を付加 |
| 条件付き書式 | `-IncludeConditionalFormats` | `payload.metadata.conditional_formats` | チャンク範囲に重なる条件付き書式を付加 |

## 一部だけ取得できる情報

| 分類 | 現状 | 取れる内容 | 取れない / 弱い内容 |
| --- | --- | --- | --- |
| 書式 | 一部のみ | 色、罫線、折り返し、整列 | フォント名、サイズ、太字、斜体、テーマ、条件付き書式反映後の見た目 |
| コメント | 一部のみ | コメント本文、threaded comment の本文と返信一覧 | 解決状態、共同編集メタデータ、完全なスレッド状態 |
| ハイパーリンク | 一部のみ | リンク先、表示文字 | ScreenTip、細かい表示状態 |
| 数式 | かなり取れる | `formula`, `formula2`, 計算結果, 表示値 | 依存関係グラフ、再計算順序、名前定義を解決した意味構造 |
| 表示状態 | 一部のみ | hidden 行列、freeze panes | ズーム、ウィンドウ位置、表示モード全体 |

## 取得しない情報

| 分類 | 取得有無 | 補足 |
| --- | --- | --- |
| 画像 | 取得しない | セル外オブジェクトとして無視する |
| 図形、SmartArt、テキストボックス | 取得しない | 図形内テキストも対象外 |
| グラフ | 取得しない | 系列、軸、凡例、ラベルなども対象外 |
| OLE オブジェクト、埋め込みファイル | 取得しない | PDF 埋め込みなども対象外 |
| フォーム / ActiveX コントロール | 取得しない | ボタン、チェックボックス、ドロップダウン部品など |
| VBA 本体 | 取得しない | `has_vba` だけ保持し、コード本体は取らない |
| Power Query / 外部接続 | 取得しない | クエリ定義、接続文字列などは対象外 |
| ピボット定義 | 取得しない | ピボットの構造やキャッシュは対象外 |
| テーブル定義 (`ListObject`) | 取得しない | 現時点ではセル値だけ扱う |
| 名前定義以外のブック機能詳細 | 取得しない | 保護、印刷設定、テーマ、スライサーなど |

## 実務上の見方

| 目的 | 向いているか | 理由 |
| --- | --- | --- |
| 数式レビュー | 向いている | 値、表示値、数式、名前定義を組み合わせて見られる |
| 表の内容分析 | 向いている | セル本体とシート構造を保持する |
| 入力ルールの確認 | 向いている | `-CollectDataValidations` で入力規則も見られる |
| 異常値の強調ルール確認 | 向いている | `-CollectConditionalFormats` でルール自体を見られる |
| Excel の見た目完全再現 | 向いていない | 画像、図形、詳細書式、条件付き書式後の見た目は完全保持しない |
| グラフ・図の解析 | 向いていない | グラフや図形は抽出対象外 |

## 使い分けの目安

| やりたいこと | おすすめコマンド |
| --- | --- |
| 普通に LLM 用ファイルを作る | `Excel2LLM.bat "C:\path\to\book.xlsx"` |
| 色や罫線も補助的に持ちたい | `Excel2LLM.bat -Extract "C:\path\to\book.xlsx" -CollectStyles` |
| 数式の意味をより読みやすくしたい | `Excel2LLM.bat -Extract "C:\path\to\book.xlsx" -CollectNamedRanges` |
| 入力候補や入力制限も分析したい | `Excel2LLM.bat -Extract "C:\path\to\book.xlsx" -CollectDataValidations` |
| 赤セル判定などのルールも分析したい | `Excel2LLM.bat -Extract "C:\path\to\book.xlsx" -CollectConditionalFormats` |
| LLM へ追加情報も渡したい | `Excel2LLM.bat -Pack "...\workbook.json" -IncludeNamedRanges -IncludeDataValidations -IncludeConditionalFormats` |
