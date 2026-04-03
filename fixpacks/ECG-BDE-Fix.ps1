[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('UN1','UN2','UN3','CUSTOM')]
    [string]$Profile = 'UN1',
    [string]$CustomDbPath = '',
    [string]$CustomNetDir = '',
    [ValidateSet('DIAG','NETDIR','HW_CAMINHO_DB','DIRECTORIES','ALL')]
    [string]$TaskMode = 'ALL',
    [ValidateSet('User','Machine','Process')]
    [string]$HwScope = 'User',
    [string]$ProfilesFile = '',
    [string]$DiagnosticScript = '',
    [string]$OutputRoot = 'C:\ECG\Output\Fixes'
)

$scriptPath = Join-Path $PSScriptRoot 'BDE-Fix-Core.ps1'
$invoke = @{
    Profile          = $Profile
    CustomDbPath     = $CustomDbPath
    CustomNetDir     = $CustomNetDir
    TaskMode         = $TaskMode
    HwScope          = $HwScope
    ProfilesFile     = $ProfilesFile
    DiagnosticScript = $DiagnosticScript
    OutputRoot       = $OutputRoot
}

& $scriptPath @invoke -WhatIf:$WhatIfPreference
exit $LASTEXITCODE
