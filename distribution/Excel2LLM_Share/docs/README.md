# Excel2LLM ドキュメント案内

- 作成日: 2026-03-12 00:50 JST
- 作成者: Codex (GPT-5)

## この文書の目的

`docs` 配下の文書が増えてきたため、用途ごとに迷わず辿れるように整理した案内ページです。

## 構成

- `guides/`
  - 利用者向けの手順書、使い方、配布方法
- `reference/`
  - データ形式、プロンプト、VBA 補助などの参照資料
- `maintainers/`
  - テスト、運用、保守者向けの実務手順
- `reports/`
  - 検証結果、レビュー結果、受け入れレポート

## まず読む順番

初めて使う人:

1. `guides/MANUAL.md`
2. `guides/USER_GUIDE.md`
3. `reference/LLM_PROMPT_FORMATS.md`

他の人へ渡したい人:

1. `guides/SHARE_PACKAGE.md`
2. `guides/MANUAL.md`

データ形式を確認したい人:

1. `reference/FORMAT.md`
2. `reference/VBA_HELPER.md`

運用や保守を担当する人:

1. `maintainers/OPERATIONS.md`
2. `reports/SECURITY_FILE_OPS_REVIEW_20260312.md`

## 文書一覧

- `guides/MANUAL.md`: 初めて使う人向けの最優先マニュアル
- `guides/USER_GUIDE.md`: 詳しめの手順書
- `guides/SHARE_PACKAGE.md`: 配布用フォルダの作り方と注意点
- `guides/USE_CASES.md`: 実務での使い方と活用例
- `reference/FORMAT.md`: `workbook.json` などのデータ構造
- `reference/LLM_PROMPT_FORMATS.md`: LLM への指示テンプレート
- `reference/VBA_HELPER.md`: VBA 補助の使い方
- `maintainers/OPERATIONS.md`: テスト、運用、リリースの実務手順
- `reports/DOMAIN_SCENARIO_REPORT_20260311.md`: ドメインシナリオ検証レポート
- `reports/SECURITY_FILE_OPS_REVIEW_20260312.md`: セキュリティ・ファイル操作レビュー
