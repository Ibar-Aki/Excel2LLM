[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('Excel2LLM', 'run_all', 'run_extract', 'run_macro_extract', 'run_pack', 'run_preflight', 'run_prompt_bundle', 'run_rebuild', 'run_verify')]
    [string]$CommandName
)

$usageMap = @{
    Excel2LLM = @(
        '使い方: Excel2LLM.bat "C:\path\to\book.xlsx" [run_all のオプション]'
        '       Excel2LLM.bat -Extract "C:\path\to\book.xlsx" [extract のオプション]'
        '       Excel2LLM.bat -Verify "C:\path\to\book.xlsx" [verify のオプション]'
        '       Excel2LLM.bat -Preflight "C:\path\to\book.xlsx" [preflight のオプション]'
        '       Excel2LLM.bat -MacroExtract "C:\path\to\book.xlsm"'
        '       Excel2LLM.bat -Pack "output\<実行結果フォルダ>\workbook.json" [pack のオプション]'
        '       Excel2LLM.bat -Rebuild "output\<実行結果フォルダ>\workbook.json" [rebuild のオプション]'
        '       Excel2LLM.bat -PromptBundle [オプション]'
        '       Excel2LLM.bat -SelfTest'
        ''
        '主な使い方:'
        '  - Excel ファイルをドラッグアンドドロップすると、そのまま一括実行します'
        '  - 引数なしで開くと、verify、復元、style 抽出も選べるメニューを表示します'
        '  - -CollectStyles や名前定義 / 入力規則 / 条件付き書式の抽出もここから呼べます'
        '  - -MacroExtract で VBA ソース抽出と LLM 用ファイル作成を行えます'
        '  - 実行後は画面を閉じず、結果を確認してからキーを押して閉じます'
        '  - 自動実行で停止させたくない場合は -NoPause を付けます'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_all = @(
        '使い方: tools\user\run_all.bat "C:\path\to\book.xlsx" [-Verify] [extract のオプション]'
        ''
        '主なオプション:'
        '  -Verify                検証も行います'
        '  -OutputDir "C:\path\to\output"  出力先フォルダを変更します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -Sheets Summary,Calc   対象シートだけを抽出します'
        '  -CollectNamedRanges    名前定義・名前付き範囲も抽出して LLM 用に含めます'
        '  -CollectDataValidations  入力規則も抽出して LLM 用に含めます'
        '  -CollectConditionalFormats  条件付き書式も抽出して LLM 用に含めます'
        ''
        '補足:'
        '  - run_all は run_extract と同じ必須 preflight（事前チェック）を通ります'
        '  - 重すぎる Excel や破損疑いのある Excel は Excel 起動前に停止します'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_extract = @(
        '使い方: tools\advanced\run_extract.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -CollectStyles         色や罫線などの補助情報も取得します'
        '  -CollectNamedRanges    名前定義・名前付き範囲も取得します'
        '  -CollectDataValidations  入力規則も取得します'
        '  -CollectConditionalFormats  条件付き書式も取得します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -Sheets Summary,Calc   指定したシートだけを抽出します'
        '  -ExcludeSheets WideTable  指定したシートを除外します'
        ''
        '補足:'
        '  - 抽出の前に必須の preflight（事前チェック）が走ります'
        '  - 重すぎる Excel や破損疑いのある Excel は Excel 起動前に停止します'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_macro_extract = @(
        '使い方: tools\advanced\run_macro_extract.bat "C:\path\to\book.xlsm" [オプション]'
        ''
        '主なオプション:'
        '  -OutputDir "C:\path\to\output"  出力先フォルダを変更します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        ''
        '補足:'
        '  - VBA 本体の可読ソースを抽出し、LLM 用の JSONL も作成します'
        '  - Excel の VBA プロジェクトアクセスが拒否される環境では raw 保存のみになる場合があります'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_pack = @(
        '使い方: tools\advanced\run_pack.bat "output\<実行結果フォルダ>\workbook.json" [オプション]'
        ''
        '主なオプション:'
        '  -ChunkBy sheet         行のまとまりを保ちやすく分割します'
        '  -ChunkBy range -MaxCells 300  セル数優先で細かく分割します'
        '  -IncludeStyles         styles.json の補助情報も含めます'
        '  -IncludeNamedRanges    名前定義・名前付き範囲も含めます'
        '  -IncludeDataValidations  入力規則も含めます'
        '  -IncludeConditionalFormats  条件付き書式も含めます'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_preflight = @(
        '使い方: tools\advanced\run_preflight.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -OutputDir "C:\path\to\output"  レポートの出力先を変更します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_prompt_bundle = @(
        '使い方: tools\user\run_prompt_bundle.bat -Scenario general [オプション]'
        ''
        '主なオプション:'
        '  -Scenario general|mechanical|accounting  用途別の指示文を選びます'
        '  -WorkbookJsonPath "output\<実行結果フォルダ>\workbook.json"   元データの保存ファイル'
        '  -JsonlPath "output\<実行結果フォルダ>\llm_package.jsonl"      LLM 用に分割したファイル'
        '  -OutputDir "output\<実行結果フォルダ>\prompt_bundle"          prompt の出力先'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_rebuild = @(
        '使い方: tools\advanced\run_rebuild.bat "output\<実行結果フォルダ>\workbook.json" [オプション]'
        ''
        '主なオプション:'
        '  -StylesJsonPath "output\<実行結果フォルダ>\styles.json"  補助情報も反映したいときに指定します'
        '  -OutputPath "C:\path\to\rebuilt.xlsx"  出力先の Excel ファイル'
        '  -Overwrite             既存の出力先を上書きします'
        ''
        '詳細: GETTING_STARTED.md'
    )
    run_verify = @(
        '使い方: tools\advanced\run_verify.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -WorkbookJsonPath "output\<実行結果フォルダ>\workbook.json"  比較対象の workbook.json を指定します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -AllowWorkbookMacros   必要な場合だけマクロ無効化を解除します'
        ''
        '詳細: GETTING_STARTED.md'
    )
}

foreach ($line in $usageMap[$CommandName]) {
    Write-Host $line
}
