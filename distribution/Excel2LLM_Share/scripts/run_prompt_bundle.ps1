[CmdletBinding()]
param(
    [string]$WorkbookJsonPath,
    [string]$JsonlPath,
    [ValidateSet('general', 'mechanical', 'accounting')]
    [string]$Scenario = 'general',
    [string]$OutputDir,
    [int]$MaxChunkPrompts = 3,
    [switch]$RedactPaths
)

. (Join-Path $PSScriptRoot 'common.ps1')

try {
    $projectRoot = Get-ProjectRoot
    if (-not $WorkbookJsonPath) {
        $WorkbookJsonPath = Join-Path (Join-Path $projectRoot 'output') 'workbook.json'
    }
    if (-not $JsonlPath) {
        $JsonlPath = Join-Path (Join-Path $projectRoot 'output') 'llm_package.jsonl'
    }
    if (-not $OutputDir) {
        $OutputDir = Join-Path (Join-Path $projectRoot 'output') 'prompt_bundle'
    }

    & (Join-Path $PSScriptRoot 'export_prompt_bundle.ps1') `
        -WorkbookJsonPath $WorkbookJsonPath `
        -JsonlPath $JsonlPath `
        -Scenario $Scenario `
        -OutputDir $OutputDir `
        -MaxChunkPrompts $MaxChunkPrompts `
        -RedactPaths:$RedactPaths

    Write-Host '=== 指示文セット作成結果 ==='
    Write-Host ('  シナリオ:   {0}' -f $Scenario)
    Write-Host ('  出力先:     {0}' -f (Get-NormalizedFullPath -Path $OutputDir))
    Write-NextStepBlock -Steps @(
        ('prompt_*.txt を開いて LLM に貼り付ける'),
        ('テンプレート確認: docs\reference\LLM_PROMPT_FORMATS.md')
    )
}
catch {
    Write-ErrorRecoverySteps -CommandName 'run_prompt_bundle'
    throw "run_prompt_bundle.ps1 の $($_.InvocationInfo.ScriptLineNumber) 行目: $($_.Exception.Message)"
}
