[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [ValidateSet('run_all', 'run_extract', 'run_pack', 'run_preflight', 'run_prompt_bundle', 'run_rebuild', 'run_verify')]
    [string]$CommandName
)

$usageMap = @{
    run_all = @(
        '使い方: run_all.bat "C:\path\to\book.xlsx" [-Verify] [extract のオプション]'
        ''
        '主なオプション:'
        '  -Verify                検証も行います'
        '  -OutputDir "C:\path\to\output"  出力先フォルダを変更します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -Sheets Summary,Calc   対象シートだけを抽出します'
        ''
        '補足:'
        '  - run_all は run_extract と同じ必須 preflight（事前チェック）を通ります'
        '  - 重すぎる Excel や破損疑いのある Excel は Excel 起動前に停止します'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
    run_extract = @(
        '使い方: run_extract.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -CollectStyles         色や罫線などの補助情報も取得します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -Sheets Summary,Calc   指定したシートだけを抽出します'
        '  -ExcludeSheets WideTable  指定したシートを除外します'
        ''
        '補足:'
        '  - 抽出の前に必須の preflight（事前チェック）が走ります'
        '  - 重すぎる Excel や破損疑いのある Excel は Excel 起動前に停止します'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
    run_pack = @(
        '使い方: run_pack.bat "output\workbook.json" [オプション]'
        ''
        '主なオプション:'
        '  -ChunkBy sheet         行のまとまりを保ちやすく分割します'
        '  -ChunkBy range -MaxCells 300  セル数優先で細かく分割します'
        '  -IncludeStyles         styles.json の補助情報も含めます'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
    run_preflight = @(
        '使い方: run_preflight.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -OutputDir "C:\path\to\output"  レポートの出力先を変更します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
    run_prompt_bundle = @(
        '使い方: run_prompt_bundle.bat -Scenario general [オプション]'
        ''
        '主なオプション:'
        '  -Scenario general|mechanical|accounting  用途別の指示文を選びます'
        '  -WorkbookJsonPath "output\workbook.json"   元データの保存ファイル'
        '  -JsonlPath "output\llm_package.jsonl"      LLM 用に分割したファイル'
        '  -OutputDir "output\prompt_bundle"          prompt の出力先'
        ''
        '詳細: docs\guides\USER_GUIDE.md'
    )
    run_rebuild = @(
        '使い方: run_rebuild.bat "output\workbook.json" [オプション]'
        ''
        '主なオプション:'
        '  -StylesJsonPath "output\styles.json"  補助情報も反映したいときに指定します'
        '  -OutputPath "C:\path\to\rebuilt.xlsx"  出力先の Excel ファイル'
        '  -Overwrite             既存の出力先を上書きします'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
    run_verify = @(
        '使い方: run_verify.bat "C:\path\to\book.xlsx" [オプション]'
        ''
        '主なオプション:'
        '  -WorkbookJsonPath "output\workbook.json"  比較対象の workbook.json を指定します'
        '  -RedactPaths           出力に絶対パスを残しにくくします'
        '  -AllowWorkbookMacros   必要な場合だけマクロ無効化を解除します'
        ''
        '詳細: docs\guides\MANUAL.md または docs\guides\USER_GUIDE.md'
    )
}

foreach ($line in $usageMap[$CommandName]) {
    Write-Host $line
}
