# 運用・リリース手順

- 作成日: 2026-03-10 01:36 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## 目的

このドキュメントは、`Excel2LLM` を日常運用し、変更を確認し、Git で管理し、必要ならリモートへ公開するまでの手順をまとめたものです。

## 日常運用の基本フロー

### 1. 変更後にテストを実行する

```bat
tools\developer\run_tests.bat
tools\user\run_self_test.bat
```

確認ポイント:

- `tools\developer\run_tests.bat` が全件成功する
- `tools\user\run_self_test.bat` が成功する
- 最新の実行結果フォルダにある `manifest.json` の `status` が `success` または意図した `warning` である

### 2. 実ファイルで動作確認する

```bat
tools\advanced\run_extract.bat "C:\Data\book.xlsx"
tools\advanced\run_pack.bat "C:\Work_Codex\Excel2LLM\output\<実行結果フォルダ>\workbook.json"
tools\advanced\run_verify.bat "C:\Data\book.xlsx" -WorkbookJsonPath "C:\Work_Codex\Excel2LLM\output\<実行結果フォルダ>\workbook.json"
```

確認ポイント:

- `workbook.json` が生成される
- `llm_package.jsonl` が生成される
- `verify_report.json` に意図しない差分がない

### 3. 差分を確認する

```bat
git status --short
git diff
```

確認ポイント:

- 生成物が誤って Git 管理対象に入っていない
- 変更内容が意図したものだけである

## リリース前チェックリスト

### 必須

- `tools\developer\run_tests.bat` が成功
- `tools\user\run_self_test.bat` が成功
- README と関連ドキュメントが最新
- 新規 `.md` に作成日と作成者が入っている
- 既存 `.md` の更新日が反映されている

### 推奨

- 実際の `.xlsx` と `.xlsm` で 1 回ずつ確認
- `-CollectStyles` あり・なしの両方を必要範囲で確認
- `ChunkBy sheet` と `ChunkBy range` の両方で確認

## コミット手順

```bat
git add .
git commit -m "Describe the change"
```

コミットメッセージの例:

- `Add release operations guide`
- `Fix sheet chunking behavior`
- `Add Pester regression tests`

## リモート未設定時の対応

このリポジトリは、初期状態では Git リモート未設定でも問題ありません。外部保存や共有が必要になった時点で設定します。

### リモートを追加する

```bat
git remote add origin <REMOTE_URL>
git remote -v
```

例:

```bat
git remote add origin https://github.com/your-org/Excel2LLM.git
```

### 初回 push

```bat
git push -u origin main
```

### ブランチ運用を始める場合

```bat
git checkout -b codex/<topic>
git push -u origin codex/<topic>
```

## 問題発生時の切り分け

### テストが失敗した

優先確認:

- 失敗したのが `tools\developer\run_tests.bat` か `tools\user\run_self_test.bat` か
- 実装変更による仕様差か
- テストフィクスチャの前提が崩れていないか

### `verify_report.json` に差分が出た

優先確認:

- 再計算が必要なブックか
- 外部リンクや volatile 関数があるか
- 保存前の Excel 内容と比較していないか

### Git に不要ファイルが出てきた

優先確認:

- `output/` や `samples/` の生成物が `.gitignore` の対象か
- Excel の一時ファイル `~$*.xlsx` が混ざっていないか

## おすすめの運用ルール

- 実装変更後はまず `tools\developer\run_tests.bat`
- Excel COM に触る変更後は `tools\user\run_self_test.bat`
- 重要変更は実ファイルでも `tools\advanced\run_verify.bat`
- コミット前に README と `docs/` の更新漏れを確認
- リモートが必要になるまでは無理に push しない
