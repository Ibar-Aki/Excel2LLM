# 配布用フォルダについて

- 作成日: 2026-03-12 00:25 JST
- 作成者: Codex (GPT-5)

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

## 配布用フォルダの生成方法

次のコマンドで生成できます。

```bat
run_build_share_package.bat
```

既定では、次のフォルダが作られます。

```text
distribution\Excel2LLM_Share
```

## 配布用フォルダに入るもの

- 実行に必要な `bat`
- 実行に必要な `scripts`
- 利用者向けの `docs`
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

- `run_extract.bat`
- `run_pack.bat`
- `run_verify.bat`
- `run_rebuild.bat`

## 配布先の人に最初に読んでもらうもの

推奨順は次です。

1. `README.md`
2. `docs\MANUAL.md`
3. `docs\USER_GUIDE.md`
4. `docs\LLM_PROMPT_FORMATS.md`

## 注意

- 配布用フォルダは生成物です
- 再生成時は `distribution\Excel2LLM_Share` を作り直します
- 配布用フォルダ内で個別に編集したファイルは、再生成すると上書きされます
