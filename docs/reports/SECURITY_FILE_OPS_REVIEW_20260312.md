# Excel2LLM セキュリティ・ファイル操作レビュー

- 作成日: 2026-03-12 00:47 JST
- 作成者: Codex (GPT-5)

## 概要

`Excel2LLM` の現行実装に対して、セキュリティとファイル操作の観点でレビューを行った。

- 対象: `scripts/common.ps1`, `scripts/extract_excel.ps1`, `scripts/excel_verify.ps1`, `scripts/build_share_package.ps1`, `scripts/export_prompt_bundle.ps1`
- 確認方法: 静的レビュー、`run_tests.bat` 実行、配布用フォルダ再生成確認
- 確認結果:
  - `run_tests.bat`: `Passed: 13 Failed: 0`
  - `run_build_share_package.bat`: 成功

結論として、以前の重大リスクだった「マクロ既定有効」「配布生成の無防備な削除」「prompt 掃除の無制限削除」は改善済みである。一方で、生成物や補助レポートに残る絶対パス、ならびにマクロ無効化失敗時の無警告動作は、なお運用上の注意点として残っている。

## リスクランキング

| 順位 | リスク | 重要度 | 状態 |
| --- | --- | --- | --- |
| 1 | マクロ無効化の失敗を握りつぶすため、環境依存で安全化が効かない可能性がある | 高 | 未解決 |
| 2 | `verify_report.json` にローカル絶対パスが残り、配布・共有時に端末情報が漏れる | 中 | 未解決 |
| 3 | `prompt_bundle_manifest.json` に絶対パスが残り、共有時に作業環境情報が漏れる | 中 | 未解決 |
| 4 | 配布フォルダ生成は override 指定時に配下外ディレクトリも再帰削除できる | 中 | 許容設計だが要注意 |
| 5 | `-RedactPaths` は抽出系だけに効き、関連出力まで一貫していない | 低 | 未解決 |

## 詳細評価

### 1. マクロ無効化の失敗を握りつぶす

- 対象: `scripts/common.ps1`
- 該当: `AutomationSecurity = 3` 設定失敗時に `catch {}` で無視している
- 参照:
  - `scripts/common.ps1:170`
  - `scripts/common.ps1:181`

評価:

- 既定で `AutomationSecurity = 3` を設定する設計自体は正しい。
- ただし設定失敗が warning にも error にもならないため、Excel 環境差分や COM 制約で無効化できなかった場合に、利用者は安全に開けたと誤認する可能性がある。
- とくに `.xlsm` を扱う業務フローでは、ここは「静かに失敗してはいけない」箇所である。

推奨:

- 少なくとも warning を `manifest.json` かコンソールへ出す。
- 可能なら「安全化できなければ失敗」とする strict モードを追加する。

### 2. `verify_report.json` に絶対パスが残る

- 対象: `scripts/excel_verify.ps1`
- 該当: `workbook_path`, `workbook_json_path`
- 参照:
  - `scripts/excel_verify.ps1:104`
  - `scripts/excel_verify.ps1:105`

評価:

- `extract_excel.ps1` には `-RedactPaths` があるが、`verify_report.json` には同等の秘匿機構がない。
- そのため、抽出物は秘匿しても、検証レポートを添付した時点で `C:\Users\...` や作業フォルダ構造が漏れる。
- 社内展開でも、ユーザー名や端末構成が混ざるのは避けたい。

推奨:

- `excel_verify.ps1` にも `-RedactPaths` を追加する。
- 少なくとも `verify_report.json` にはファイル名のみ、または相対識別子だけを残す。

### 3. `prompt_bundle_manifest.json` に絶対パスが残る

- 対象: `scripts/export_prompt_bundle.ps1`
- 該当: `workbook_json_path`, `jsonl_path`, `prompts[].path`
- 参照:
  - `scripts/export_prompt_bundle.ps1:181`
  - `scripts/export_prompt_bundle.ps1:193`
  - `scripts/export_prompt_bundle.ps1:194`

評価:

- 削除ガードは強化されているが、manifest 自体には依然として絶対パスが残る。
- prompt 束は他者へ渡す運用になりやすいため、むしろこの manifest の方が漏えい面では実害が出やすい。
- 共有物に `C:\Users\AKIHIRO\...` が残る設計は避けた方がよい。

推奨:

- `prompts[].path` は `file_name` か `relative_path` に置き換える。
- `workbook_json_path` / `jsonl_path` も basename か relative path に落とす。

### 4. 配布フォルダ生成は override 指定時に広い削除権限を持つ

- 対象: `scripts/build_share_package.ps1`
- 該当: `Remove-Item -Recurse -Force`
- 参照:
  - `scripts/build_share_package.ps1:21`
  - `scripts/build_share_package.ps1:26`
  - `scripts/build_share_package.ps1:30`

評価:

- 既定では `distribution\` 配下に限定されており、通常運用の安全性は改善済み。
- ただし `-AllowOutsideDistribution -ForceCleanOutputDir` を併用すると、任意の既存ディレクトリを再帰削除できる。
- これは意図された escape hatch だが、運用ドキュメントを読まずに流用すると事故余地は残る。

推奨:

- 現状のままでも実用上は許容できる。
- ただし業務配布用なら、override 時に確認メッセージを追加するか、より明示的なフラグ名へ変更した方がよい。

### 5. `-RedactPaths` の効き方が出力ごとに揃っていない

- 対象: `scripts/extract_excel.ps1`, `scripts/excel_verify.ps1`, `scripts/export_prompt_bundle.ps1`
- 参照:
  - `scripts/extract_excel.ps1:230`
  - `scripts/extract_excel.ps1:254`
  - `scripts/excel_verify.ps1:104`
  - `scripts/export_prompt_bundle.ps1:193`

評価:

- 抽出物では path 秘匿ができる一方、検証レポートと prompt manifest には残る。
- そのため利用者視点では「秘匿オプションを付けたのに、まだ漏れる場所がある」状態で、一貫性が弱い。

推奨:

- path 秘匿ポリシーを共通化する。
- `extract`, `verify`, `prompt export` で同じオプション名と同じ挙動を提供する。

## 良い点

- Excel 起動時に既定でマクロ無効化を試みる設計へ改善済み。
- 配布フォルダ生成は既定で `distribution\` 配下以外の削除を拒否する。
- prompt 再生成は manifest 改ざんによる配下外削除を防ぐ。
- `share_manifest.json` から配布元の絶対パスは除去済み。
- 回帰テストで、path 秘匿、prompt cleanup guard、share package guard が固定されている。

## 総評

現状の `Excel2LLM` は、「信頼済み Excel を社内ローカル環境で扱う」前提なら、以前よりかなり安全になっている。特に破壊的削除と prompt cleanup の暴走は抑えられている。

一方で、「共有物にローカル情報を残さない」「安全化が効かなかったら明示的に分かる」という観点では、まだ hardening の余地がある。次に手を入れるなら、優先順は次の通りである。

1. `AutomationSecurity` 設定失敗の warning / fail-fast 化
2. `excel_verify.ps1` の `-RedactPaths` 対応
3. `export_prompt_bundle.ps1` の相対パス化
