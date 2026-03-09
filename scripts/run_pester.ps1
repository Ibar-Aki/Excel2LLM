[CmdletBinding()]
param(
    [string]$Path = (Join-Path (Split-Path -Path $PSScriptRoot -Parent) 'tests')
)

$module = Get-Module -ListAvailable Pester | Sort-Object Version -Descending | Select-Object -First 1
if ($null -eq $module) {
    throw 'Pester module is not available on this machine.'
}

Import-Module $module.Path -Force
Invoke-Pester -Path $Path -EnableExit
