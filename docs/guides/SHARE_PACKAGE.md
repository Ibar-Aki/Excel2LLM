# 配布用フォルダについて

- 作成日: 2026-03-12 00:25 JST
- 作成者: Codex (GPT-5)
- 更新日: 2026-03-28

## この文書の目的

この文書は、`Excel2LLM` を他の人へ渡すための配布用フォルダについて説明するためのものです。

## 配布用フォルダの考え方

普段の開発フォルダには、次のものが混ざります。

- 検証用の `output`
- 一時的なサンプル Excel
- テストコード
- ドメイン検証レポート

これらは開発や検証には必要ですが、配布先の利用者には必ずしも必要ではありません。

そのため、配布時は必要な実行ファイル、スクリプト、最重要ドキュメントだけをまとめた専用フォルダを使います。
利用者向けの手順は `GETTING_STARTED.md` に一本化し、配布用フォルダには不要なガイドを入れない方針です。

## 配布用フォルダの生成方法

次のコマンドで生成できます。

```bat
tools\developer\run_build_share_package.bat
```

既定では、次のフォルダが作られます。

```text
distribution\Excel2LLM_Share
```

`distribution\` 配下以外へ生成したい場合は、明示フラグが必要です。

```bat
tools\developer\run_build_share_package.bat -OutputDir "C:\Temp\Excel2LLM_Share" -AllowOutsideDistribution -ForceCleanOutputDir
```

## 配布用フォルダに入るもの

- 実行に必要な `bat`
- 実行に必要な `scripts`
- 利用者向けの `docs`
- 配布先の最初の案内 `GETTING_STARTED.md`
- `templates`
- 空の `output`
- 空の `samples`

## 配布用フォルダに入れないもの

- `tests`
- 開発中の一時出力
- 過去の検証用 `output` 配下の生成物
- Git 履歴そのもの

## 利用者への渡し方

基本は、`distribution\Excel2LLM_Share` フォルダごと渡せば十分です。

配布先では、そのフォルダの中で次を使えば動作します。

- `Excel2LLM.bat`
- `tools\user\run_all.bat`
- `tools\user\run_prompt_bundle.bat`
- `tools\user\run_self_test.bat`
- `tools\advanced\run_extract.bat`
- `tools\advanced\run_preflight.bat`
- `tools\advanced\run_pack.bat`
- `tools\advanced\run_verify.bat`
- `tools\advanced\run_rebuild.bat`

## 配布先の人に最初に読んでもらうもの

推奨順は次です。

1. `README.md`
2. `GETTING_STARTED.md`
3. `docs\reference\LLM_PROMPT_FORMATS.md`

役割の違い:

- `README.md`
  - このフォルダで何ができるかを最初に把握する入口
- `GETTING_STARTED.md`
  - 利用者向けの手順をまとめた唯一の案内

## 注意

- 配布用フォルダは生成物です
- 再生成時は `distribution\Excel2LLM_Share` を作り直します
- 配布用フォルダ内で個別に編集したファイルは、再生成すると上書きされます
- 配下外の既存ディレクトリは、明示フラグなしでは削除しません
- `share_manifest.json` には配布元 PC の絶対パスを残さないようにしています
- `Excel2LLM.bat -Preflight` を使うと、本番抽出の前に危険ファイルを止められます
