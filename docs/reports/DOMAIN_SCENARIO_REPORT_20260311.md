# Excel2LLM ドメインシナリオ検証レポート

- 作成日: 2026-03-11 08:05 JST
- 作成者: Codex (GPT-5)

## 概要

機械設計向けと会計向けの 2 系統について、次の流れを一気通貫で実施しました。

1. Excel サンプルの作成
2. `Excel2LLM` による抽出
3. `llm_package.jsonl` とプロンプト束の生成
4. LLM 観点での問題抽出と改善方針整理
5. 改善版 Excel の作成
6. 再抽出、再検証、比較
7. テスト、レビュー、修正

今回の検証は、単にサンプルを作るだけでなく、改善前後を同条件で比較し、再現可能な形で出力を残すことを目的にしています。

## 実施日時

- 実施日: 2026-03-11
- 主要な自動実行開始: 2026-03-11 07:52 JST
- 主要な自動実行終了: 2026-03-11 07:54 JST

## 対象環境

- OS: Windows
- Excel: M365 Excel
- 実行基盤: PowerShell + Excel COM + bat
- 対象機能:
  - `extract_excel.ps1`
  - `pack_for_llm.ps1`
  - `excel_verify.ps1`
  - `rebuild_excel.ps1`
  - 新規追加のドメインサンプル生成、プロンプト束生成、受け入れ実行

## 今回追加した補助機能

今回の検証で、ユーザー負担を減らすために次を追加実装しました。

- [create_domain_sample_workbooks.ps1](/C:/Work_Codex/Excel2LLM/scripts/create_domain_sample_workbooks.ps1)
  - 機械設計向け、会計向けのサンプル Excel を改善前後 2 パターンで自動生成
- [export_prompt_bundle.ps1](/C:/Work_Codex/Excel2LLM/scripts/export_prompt_bundle.ps1)
  - `llm_package.jsonl` から、用途別のコピペ用プロンプト束を自動生成
- [run_domain_acceptance.ps1](/C:/Work_Codex/Excel2LLM/scripts/run_domain_acceptance.ps1)
  - サンプル生成、抽出、pack、verify、プロンプト出力、集計を一括実行
- [run_domain_acceptance.bat](/C:/Work_Codex/Excel2LLM/run_domain_acceptance.bat)
  - 上記の実行入口

## 実行シナリオ

### シナリオ 1: 機械設計向け

- 元版ファイル: `mechanical_original.xlsx`
- 改善版ファイル: `mechanical_improved.xlsx`
- 主な対象シート:
  - 元版: `Inputs`, `Calc`, `Review`
  - 改善版: `Inputs`, `ShaftSizing`, `Checks`
- 想定業務:
  - 軸径の簡易計算シートを読み解く
  - 計算手順と説明性を改善する

### シナリオ 2: 会計向け

- 元版ファイル: `accounting_original.xlsx`
- 改善版ファイル: `accounting_improved.xlsx`
- 主な対象シート:
  - 元版: `Transactions`, `Budget`, `Summary`, `Notes`
  - 改善版: `Transactions`, `Budget`, `Summary`, `Checks`
- 想定業務:
  - 予実差異の見える化
  - 数値確認と集計表の改善

## 定量結果

出力元: [scenario_summary.json](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/scenario_summary.json)

### 機械設計向け

- 元版
  - シート数: 3
  - セル数: 66
  - 数式数: 7
  - 結合範囲数: 1
  - chunk 数: 3
  - prompt 数: 3
  - verify mismatch: 0
  - 所要時間: 3.74 秒
- 改善版
  - シート数: 3
  - セル数: 68
  - 数式数: 8
  - 結合範囲数: 0
  - chunk 数: 3
  - prompt 数: 3
  - verify mismatch: 0
  - 所要時間: 3.86 秒

### 会計向け

- 元版
  - シート数: 4
  - セル数: 134
  - 数式数: 18
  - 結合範囲数: 1
  - chunk 数: 4
  - prompt 数: 3
  - verify mismatch: 0
  - 所要時間: 4.50 秒
- 改善版
  - シート数: 4
  - セル数: 138
  - 数式数: 32
  - 結合範囲数: 0
  - chunk 数: 4
  - prompt 数: 3
  - verify mismatch: 0
  - 所要時間: 5.12 秒

## LLM プロンプトのテストパターン

### 機械設計向け

- パターン A: 計算手順レビュー
  - 入力: [01_mechanical_Calc.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/mechanical_original/prompts/01_mechanical_Calc.txt)
  - 目的: 計算の流れ、主要数式、冗長計算の検出
- パターン B: 入力前提レビュー
  - 入力: [02_mechanical_Inputs.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/mechanical_original/prompts/02_mechanical_Inputs.txt)
  - 目的: 単位、前提条件、入力値の意味の整理
- パターン C: チェック状態レビュー
  - 入力: [03_mechanical_Review.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/mechanical_original/prompts/03_mechanical_Review.txt)
  - 目的: 未レビュー箇所の特定

### 会計向け

- パターン A: サマリー表レビュー
  - 入力: [01_accounting_Summary.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/accounting_original/prompts/01_accounting_Summary.txt)
  - 目的: 売上、費用、利益、予実差異の把握
- パターン B: 取引明細レビュー
  - 入力: [02_accounting_Transactions.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/accounting_original/prompts/02_accounting_Transactions.txt)
  - 目的: 明細の型、区分、異常値候補の把握
- パターン C: 予算表レビュー
  - 入力: [03_accounting_Budget.txt](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311/accounting_original/prompts/03_accounting_Budget.txt)
  - 目的: 参照元予算データの確認

## LLM 改善テストの結果

### 機械設計向けの主な発見

元版 `Calc` シートでは、次が問題として明確でした。

- `B4` と `B9` に必要軸径の同じ式が重複していた
- `ReqDia`, `ExistingDia`, `Judge`, `Tmp` のように、意味が省略されていて説明性が低かった
- `B2 = Inputs!B2*1000` の単位換算理由がセル上で説明されていなかった
- 判定結果は `PASS/FAIL` のみで、改善行動につながるメッセージがなかった
- `Review` シートで「未レビュー」が見えているが、計算シート本体には反映されていなかった

#### 元版の主要セル

- 必要軸径: `Calc!B4 = 34.27469`
- 安全比: `Calc!B6 = 1.167042`
- 判定: `Calc!B7 = PASS`
- 重複式: `Calc!B9 = 34.27469`
- 仮サービス係数: `Inputs!B4 = 1.6`

#### 改善内容

- `Calc` を `ShaftSizing` に改名
- 説明列 `Explanation` を追加
- `RequiredDiameter_mm`、`ExistingDiameter_mm` などの明示的な行名へ変更
- `Margin_mm` と `Recommendation` を追加
- 重複式 `B9` を廃止
- `Checks` シートでレビュー項目を明示
- 仮置きだったサービス係数をレビュー済み値 `1.5` に更新したサンプルに変更

#### 改善後の主要セル

- 必要軸径: `ShaftSizing!B4 = 33.6`
- 安全比: `ShaftSizing!B6 = 1.19`
- 余裕量: `ShaftSizing!B7 = 6.4 mm`
- 推奨: `ShaftSizing!B8 = Current diameter acceptable`
- レビュー済みサービス係数: `Inputs!B4 = 1.5`

#### 判断

- LLM による説明、レビュー、改善提案の対象として、元版は十分に問題を含んでいた
- 改善版では、計算根拠、判断結果、レビュー状態が追跡しやすくなった
- ユーザーが「この式は何か」「次に何をすべきか」を読み取りやすくなった

### 会計向けの主な発見

元版 `Summary` シートでは、次が問題として明確でした。

- 売上、費用、利益まではあるが、予実差異がセルとして見えない
- `G列` が `Memo` 固定文言で、分析列として機能していない
- 利益率が見えない
- `Notes` シートで未確認状態が残っている
- 表の末尾にレビュー用の結合セルがあり、確認対象と計算対象が混ざっていた

#### 元版の主要セル

- Design 売上: `Summary!B2 = 180000`
- Design 費用: `Summary!C2 = 90000`
- Design 利益: `Summary!D2 = 90000`
- Design 予算売上: `Summary!E2 = 210000`
- Design 予算費用: `Summary!F2 = 95000`
- `Summary!G2 = "Need variance view"`
- 合計売上: `Summary!B6 = 520000`
- 合計費用: `Summary!C6 = 253000`
- 合計利益: `Summary!D6 = 267000`

#### 改善内容

- `Summary` に `RevenueVariance`, `CostVariance`, `ProfitMargin` を追加
- 合計行にも差異と利益率を追加
- `Notes` を `Checks` に変更し、未確認メモをチェック観点へ変換
- レビュー用の結合セルを廃止
- 数式セルを増やし、判断根拠を表内で追えるようにした

#### 改善後の主要セル

- Design 売上差異: `Summary!G2 = -30000`
- Design 費用差異: `Summary!H2 = -5000`
- Design 利益率: `Summary!I2 = 50.0%`
- 合計売上: `Summary!B6 = 520000`
- 合計利益: `Summary!D6 = 267000`
- 合計利益率: `Summary!I6 = 51.3%`

#### 判断

- 元版は「数字はあるが判断材料が足りない」状態だった
- 改善版では、月次レビューや予実確認に必要な差異情報が直接見える
- 会計担当者が別計算なしで異常箇所候補に到達しやすくなった

## レビューで見つかった実装上の問題と修正

今回の実装中に、次の問題を見つけて修正しました。

### 1. ドメインサンプル生成で COM 型変換が不安定

- 症状:
  - 配列要素や `if` 式の結果をそのまま `Value2` に代入すると、`System.Double` や `System.Int32` のキャスト例外が起きた
- 修正:
  - `Set-CellValue` ヘルパーを追加し、型を明示的に変換して書き込むように変更
- 対象:
  - [create_domain_sample_workbooks.ps1](/C:/Work_Codex/Excel2LLM/scripts/create_domain_sample_workbooks.ps1)

### 2. 受け入れ実行スクリプトの verify 呼び出し引数が誤っていた

- 症状:
  - `excel_verify.ps1` に存在しない `-OutputPath` を渡して失敗した
- 修正:
  - `-OutputDir` に修正
- 対象:
  - [run_domain_acceptance.ps1](/C:/Work_Codex/Excel2LLM/scripts/run_domain_acceptance.ps1)

### 3. プロンプト束の優先順位が悪く、重要シートが先頭に来ない

- 症状:
  - 機械設計では `Calc` より `Review`
  - 会計では `Summary` より `Notes`
  - が先頭に出ることがあり、ユーザーが毎回探す必要があった
- 修正:
  - `formula_cells` 数の多い chunk を優先して prompt を並べるよう変更
- 対象:
  - [export_prompt_bundle.ps1](/C:/Work_Codex/Excel2LLM/scripts/export_prompt_bundle.ps1)

## テスト結果

### 統合テスト

- 実行コマンド: `run_tests.bat`
- 結果: `Passed: 10 Failed: 0`
- 所要時間: `161` 秒

### ドメイン受け入れ実行

- 実行コマンド: `run_domain_acceptance.bat`
- 出力先: [output/domain_acceptance_20260311](/C:/Work_Codex/Excel2LLM/output/domain_acceptance_20260311)
- 結果概要:
  - 4 シナリオすべて成功
  - verify mismatch は全件 `0`
  - prompt bundle 生成成功

## ユーザー負担を減らす観点での改善確認

今回の作業で、ユーザー負担を下げる観点では次が有効でした。

- ドメイン別サンプルと改善版を自動生成できる
- 抽出から prompt 束生成まで一括で回せる
- prompt の先頭に、最も数式量の多い重要シートが来る
- 元版と改善版を同じ指標で比較できる
- レポート化の根拠となる JSON が残る

特に、prompt の優先順位修正は小さい変更ですが効果が大きく、ユーザーが毎回「どのファイルを LLM に貼るべきか」を探す手間を減らします。

## エラー有無

- 最終結果としてのエラー: なし
- 実装途中で修正したエラー:
  - COM 型変換エラー
  - verify 引数名の誤り
  - prompt 順序の不適切さ

## 失敗時の原因推定

今回修正した内容から、今後失敗しやすい箇所は次です。

- Excel COM への値代入時の型揺れ
- PowerShell での引数名ミス
- 利用者目線では重要度の低いシートが prompt の先頭に来る並び順

## 今回の結論

機械設計向け、会計向けの両方で、次を確認できました。

- `Excel2LLM` の抽出、pack、verify は安定している
- LLM 向け prompt 束を生成し、改善前後の比較に使える
- 改善版 Excel をあらかじめ設計し、効果を定量比較できる
- 実装中に見つかった不足機能は追加実装と修正で吸収できた

現時点で、今回追加したドメイン検証フローは、今後の実案件前テストの叩き台として十分使える状態です。
