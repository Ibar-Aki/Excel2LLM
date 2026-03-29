[CmdletBinding()]
param(
    [string]$ProjectRoot
)

. (Join-Path $PSScriptRoot 'common.ps1')

if (-not $ProjectRoot) {
    $ProjectRoot = Split-Path -Path $PSScriptRoot -Parent
}

$runAllScript = Join-Path $ProjectRoot 'scripts\run_all.ps1'
$extractScript = Join-Path $ProjectRoot 'scripts\extract_excel.ps1'
$packScript = Join-Path $ProjectRoot 'scripts\pack_for_llm.ps1'
$preflightScript = Join-Path $ProjectRoot 'scripts\preflight_excel.ps1'
$verifyScript = Join-Path $ProjectRoot 'scripts\excel_verify.ps1'
$rebuildScript = Join-Path $ProjectRoot 'scripts\rebuild_excel.ps1'
$promptBundleScript = Join-Path $ProjectRoot 'scripts\run_prompt_bundle.ps1'
$selfTestScript = Join-Path $ProjectRoot 'scripts\run_self_test.ps1'
$advancedDir = Join-Path $ProjectRoot 'tools\advanced'
$outputDir = Join-Path $ProjectRoot 'output'

function Read-RequiredInput {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    $value = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($value)) {
        throw '入力が空のため処理を中止しました。'
    }

    return $value.Trim()
}

function Read-OptionalInput {
    param(
        [Parameter(Mandatory)]
        [string]$Prompt
    )

    $value = Read-Host $Prompt
    if ([string]::IsNullOrWhiteSpace($value)) {
        return $null
    }

    return $value.Trim()
}

function Get-DefaultWorkbookJsonPath {
    $latestOutputDir = Get-LatestOutputDirectory
    return Join-Path $latestOutputDir 'workbook.json'
}

function Get-DefaultStylesJsonPath {
    param(
        [Parameter(Mandatory)]
        [string]$WorkbookJsonPath
    )

    $candidate = Join-Path (Split-Path -Path $WorkbookJsonPath -Parent) 'styles.json'
    if (Test-Path -LiteralPath $candidate) {
        return $candidate
    }

    return $null
}

Write-Host '=== Excel2LLM ==='
Write-Host '1. Excel を処理する'
Write-Host '2. Excel を処理して照合もする'
Write-Host '3. 見た目情報や追加ルールも含めて抽出する'
Write-Host '4. 事前チェックだけ行う'
Write-Host '5. 抽出結果と元 Excel を照合する'
Write-Host '6. 抽出結果を分割し直す'
Write-Host '7. 抽出結果から Excel を復元する'
Write-Host '8. 最新結果から指示文セットを作る'
Write-Host '9. 動作確認をする'
Write-Host '10. 詳細機能フォルダを開く'
Write-Host '0. 終了する'

$choice = Read-Host '番号を入力してください'

try {
    switch ($choice) {
        '1' {
            $excelPath = Read-RequiredInput -Prompt 'Excel ファイルのパスを入力してください'
            & $runAllScript -ExcelPath $excelPath
            exit 0
        }
        '2' {
            $excelPath = Read-RequiredInput -Prompt 'Excel ファイルのパスを入力してください'
            & $runAllScript -ExcelPath $excelPath -Verify
            exit 0
        }
        '3' {
            $excelPath = Read-RequiredInput -Prompt 'Excel ファイルのパスを入力してください'
            & $extractScript -ExcelPath $excelPath -CollectStyles -CollectNamedRanges -CollectDataValidations -CollectConditionalFormats
            exit 0
        }
        '4' {
            $excelPath = Read-RequiredInput -Prompt 'Excel ファイルのパスを入力してください'
            & $preflightScript -ExcelPath $excelPath
            exit 0
        }
        '5' {
            $excelPath = Read-RequiredInput -Prompt '元の Excel ファイルのパスを入力してください'
            $defaultWorkbookJsonPath = Get-DefaultWorkbookJsonPath
            $workbookJsonPath = Read-OptionalInput -Prompt ("workbook.json のパスを入力してください（Enter で最新結果を使用: {0}）" -f $defaultWorkbookJsonPath)
            if (-not $workbookJsonPath) {
                $workbookJsonPath = $defaultWorkbookJsonPath
            }

            & $verifyScript -ExcelPath $excelPath -WorkbookJsonPath $workbookJsonPath
            exit 0
        }
        '6' {
            $defaultWorkbookJsonPath = Get-DefaultWorkbookJsonPath
            $workbookJsonPath = Read-OptionalInput -Prompt ("workbook.json のパスを入力してください（Enter で最新結果を使用: {0}）" -f $defaultWorkbookJsonPath)
            if (-not $workbookJsonPath) {
                $workbookJsonPath = $defaultWorkbookJsonPath
            }

            & $packScript -WorkbookJsonPath $workbookJsonPath
            exit 0
        }
        '7' {
            $defaultWorkbookJsonPath = Get-DefaultWorkbookJsonPath
            $workbookJsonPath = Read-OptionalInput -Prompt ("workbook.json のパスを入力してください（Enter で最新結果を使用: {0}）" -f $defaultWorkbookJsonPath)
            if (-not $workbookJsonPath) {
                $workbookJsonPath = $defaultWorkbookJsonPath
            }

            $defaultStylesJsonPath = Get-DefaultStylesJsonPath -WorkbookJsonPath $workbookJsonPath
            $stylesPrompt = if ($defaultStylesJsonPath) {
                "styles.json のパスを入力してください（Enter で既定を使用: $defaultStylesJsonPath、不要なら空のまま Enter）"
            }
            else {
                'styles.json のパスを入力してください（不要なら空のまま Enter）'
            }
            $stylesJsonPath = Read-OptionalInput -Prompt $stylesPrompt
            if (-not $stylesJsonPath) {
                $stylesJsonPath = $defaultStylesJsonPath
            }

            if ($stylesJsonPath) {
                & $rebuildScript -WorkbookJsonPath $workbookJsonPath -StylesJsonPath $stylesJsonPath
            }
            else {
                & $rebuildScript -WorkbookJsonPath $workbookJsonPath
            }
            exit 0
        }
        '8' {
            Write-Host 'シナリオを選んでください: 1=general  2=mechanical  3=accounting'
            $scenarioChoice = Read-Host '番号を入力してください（Enter で 1）'
            $scenario = switch ($scenarioChoice) {
                '2' { 'mechanical' }
                '3' { 'accounting' }
                default { 'general' }
            }

            & $promptBundleScript -Scenario $scenario
            exit 0
        }
        '9' {
            & $selfTestScript
            exit 0
        }
        '10' {
            Start-Process explorer.exe $advancedDir | Out-Null
            exit 0
        }
        '0' {
            exit 0
        }
        default {
            Write-Host '0 から 10 の番号を入力してください。'
            exit 1
        }
    }
}
catch {
    Write-Host $_.Exception.Message
    exit 1
}
