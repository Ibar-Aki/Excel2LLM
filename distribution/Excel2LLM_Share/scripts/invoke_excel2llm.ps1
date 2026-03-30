. (Join-Path $PSScriptRoot 'common.ps1')

function Convert-ToPowerShellArgumentToken {
    param(
        [Parameter(Mandatory)]
        [string]$Value
    )

    if ($Value -match '^-{1,2}[A-Za-z]') {
        return $Value
    }

    return "'" + $Value.Replace("'", "''") + "'"
}

function Invoke-Excel2LLMTarget {
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,
        [string[]]$ScriptArguments
    )

    $tokens = [System.Collections.Generic.List[string]]::new()
    [void]$tokens.Add("& '" + $ScriptPath.Replace("'", "''") + "'")

    foreach ($argument in @($ScriptArguments)) {
        [void]$tokens.Add((Convert-ToPowerShellArgumentToken -Value ([string]$argument)))
    }

    $commandText = $tokens -join ' '
    Invoke-Expression $commandText
    exit 0
}

function Test-RequiresPositionalPath {
    param(
        [Parameter(Mandatory)]
        [string]$CommandName
    )

    return @('-runall', '-extract', '-verify', '-preflight', '-macroextract', '-pack', '-rebuild') -contains $CommandName.ToLowerInvariant()
}

$projectRoot = Get-ProjectRoot
$noPauseFlagPath = [string]$env:EXCEL2LLM_NO_PAUSE_FLAG
$noPauseRequested = $false
$effectiveArguments = New-Object System.Collections.Generic.List[string]

foreach ($argument in @($args)) {
    if ([string]$argument -eq '-NoPause') {
        $noPauseRequested = $true
        continue
    }

    [void]$effectiveArguments.Add([string]$argument)
}

if ($noPauseRequested -and -not [string]::IsNullOrWhiteSpace($noPauseFlagPath)) {
    Set-Content -LiteralPath $noPauseFlagPath -Value '1' -NoNewline
}

if ($effectiveArguments.Count -eq 0) {
    & (Join-Path $PSScriptRoot 'launch_menu.ps1') -ProjectRoot $projectRoot
    exit $LASTEXITCODE
}

$commandName = [string]$effectiveArguments[0]
$remainingArguments = @($effectiveArguments | Select-Object -Skip 1)

if ((Test-RequiresPositionalPath -CommandName $commandName) -and $remainingArguments.Count -eq 0) {
    & (Join-Path $PSScriptRoot 'show_usage.ps1') -CommandName Excel2LLM
    exit 1
}

switch -Regex ($commandName.ToLowerInvariant()) {
    '^-h$|^--help$|^/\?$' {
        & (Join-Path $PSScriptRoot 'show_usage.ps1') -CommandName Excel2LLM
        exit 1
    }
    '^-runall$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\run_all.ps1') -ScriptArguments $remainingArguments
    }
    '^-extract$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\extract_excel.ps1') -ScriptArguments $remainingArguments
    }
    '^-verify$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\excel_verify.ps1') -ScriptArguments $remainingArguments
    }
    '^-preflight$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\preflight_excel.ps1') -ScriptArguments $remainingArguments
    }
    '^-macroextract$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\macro_extract.ps1') -ScriptArguments $remainingArguments
    }
    '^-pack$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\pack_for_llm.ps1') -ScriptArguments $remainingArguments
    }
    '^-rebuild$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\rebuild_excel.ps1') -ScriptArguments $remainingArguments
    }
    '^-promptbundle$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\run_prompt_bundle.ps1') -ScriptArguments $remainingArguments
    }
    '^-selftest$' {
        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\run_self_test.ps1') -ScriptArguments $remainingArguments
    }
    default {
        if ($commandName.StartsWith('-')) {
            & (Join-Path $PSScriptRoot 'show_usage.ps1') -CommandName Excel2LLM
            exit 1
        }

        Invoke-Excel2LLMTarget -ScriptPath (Join-Path $projectRoot 'scripts\run_all.ps1') -ScriptArguments $effectiveArguments
    }
}
