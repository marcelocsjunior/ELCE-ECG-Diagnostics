<#
.SYNOPSIS
    Diagnóstico e correção controlada do contrato BDE/ECG por unidade.
.DESCRIPTION
    Mantém a correção fora do core read-only. Resolve perfil por unidade,
    valida HW_CAMINHO_DB/NETDIR e, se autorizado, aplica a correção no HKLM.
    Usa wrapper BAT/ExecutionPolicy Bypass para eliminar necessidade de comando manual.
#>

[CmdletBinding()]
param(
    [switch]$Fix,
    [switch]$CreateMissingDirs,
    [string]$Unit = 'AUTO',
    [string]$ProfileFile = '',
    [switch]$OpenRunbook
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
if ([string]::IsNullOrWhiteSpace($ProfileFile)) {
    $ProfileFile = Join-Path $ScriptRoot 'ECG_UnitProfiles.json'
}
$RunbookPath = Join-Path $ScriptRoot 'runbookECG-BDE-Fix.txt'
$OutputRoot = 'C:\ECG\Output\Fixes'

function Restart-WithBypassIfNeeded {
    if (-not $MyInvocation.MyCommand.Path) { return }
    if ($env:ECG_BDE_FIX_BYPASS_RESTARTED -eq '1') { return }

    $effectivePolicy = $null
    try { $effectivePolicy = Get-ExecutionPolicy -ErrorAction Stop } catch { $effectivePolicy = $null }

    if ($effectivePolicy -in @('Restricted','AllSigned')) {
        Write-Host "Política de execução '$effectivePolicy' detectada. Reiniciando com Bypass..." -ForegroundColor Yellow
        $args = New-Object System.Collections.Generic.List[string]
        $args.Add('-NoProfile')
        $args.Add('-ExecutionPolicy')
        $args.Add('Bypass')
        $args.Add('-File')
        $args.Add($MyInvocation.MyCommand.Path)
        if ($Fix) { $args.Add('-Fix') }
        if ($CreateMissingDirs) { $args.Add('-CreateMissingDirs') }
        if (-not [string]::IsNullOrWhiteSpace($Unit)) { $args.Add('-Unit'); $args.Add($Unit) }
        if (-not [string]::IsNullOrWhiteSpace($ProfileFile)) { $args.Add('-ProfileFile'); $args.Add($ProfileFile) }
        if ($OpenRunbook) { $args.Add('-OpenRunbook') }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = ($args | ForEach-Object {
            if ($_ -match '\s') { '"' + ($_ -replace '"', '""') + '"' } else { $_ }
        }) -join ' '
        $psi.UseShellExecute = $false
        $psi.EnvironmentVariables['ECG_BDE_FIX_BYPASS_RESTARTED'] = '1'
        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit()
        exit $proc.ExitCode
    }
}

function Test-IsAdmin {
    try {
        $id = [Security.Principal.WindowsIdentity]::GetCurrent()
        $pr = New-Object Security.Principal.WindowsPrincipal($id)
        return $pr.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    }
    catch {
        return $false
    }
}

function Restart-ElevatedIfNeeded {
    param([switch]$NeedElevation)

    if (-not $NeedElevation) { return }
    if (Test-IsAdmin) { return }
    if ($env:ECG_BDE_FIX_ELEVATED -eq '1') { return }

    Write-Host "Elevação necessária para alterar HKLM e/ou criar diretórios." -ForegroundColor Yellow
    $args = New-Object System.Collections.Generic.List[string]
    $args.Add('-NoProfile')
    $args.Add('-ExecutionPolicy')
    $args.Add('Bypass')
    $args.Add('-File')
    $args.Add($MyInvocation.MyCommand.Path)
    if ($Fix) { $args.Add('-Fix') }
    if ($CreateMissingDirs) { $args.Add('-CreateMissingDirs') }
    if (-not [string]::IsNullOrWhiteSpace($Unit)) { $args.Add('-Unit'); $args.Add($Unit) }
    if (-not [string]::IsNullOrWhiteSpace($ProfileFile)) { $args.Add('-ProfileFile'); $args.Add($ProfileFile) }
    if ($OpenRunbook) { $args.Add('-OpenRunbook') }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.Arguments = ($args | ForEach-Object {
        if ($_ -match '\s') { '"' + ($_ -replace '"', '""') + '"' } else { $_ }
    }) -join ' '
    $psi.Verb = 'runas'
    $psi.UseShellExecute = $true

    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    exit $proc.ExitCode
}

function Get-DefaultUnitProfiles {
    $json = @'
{
  "version": "2026-03-28",
  "defaultExePath": "C:\\HW\\ECG\\ECGV6.exe",
  "units": {
    "UN1": {
      "name": "Unidade 1",
      "topology": "FILE_SERVER_DEDICADO",
      "dbPath": "\\\\SRVVM1-FS01\\FS\\ECG\\HW\\Database",
      "netDirPath": "\\\\SRVVM1-FS01\\FS\\ECG\\HW\\Database\\NetDir",
      "fallbackDbPath": "P:\\ECG\\HW\\Database",
      "fileServerHost": "SRVVM1-FS01",
      "computerPatterns": ["^ELCUN1-"]
    },
    "UN2": {
      "name": "Unidade 2",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN2-ECG\\hw",
      "netDirPath": "\\\\ELCUN2-ECG\\hw\\NetDir",
      "fallbackDbPath": "",
      "fileServerHost": "",
      "computerPatterns": ["^ELCUN2-"]
    },
    "UN3": {
      "name": "Unidade 3",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN3-ECG\\hw",
      "netDirPath": "\\\\ELCUN3-ECG\\hw\\NetDir",
      "fallbackDbPath": "",
      "fileServerHost": "",
      "computerPatterns": ["^ELCUN3-"]
    }
  }
}
'@
    return ($json | ConvertFrom-Json)
}

function Get-UnitProfiles {
    $defaults = Get-DefaultUnitProfiles
    if (Test-Path -LiteralPath $ProfileFile) {
        try {
            $external = Get-Content -LiteralPath $ProfileFile -Raw | ConvertFrom-Json
            if ($null -ne $external -and $null -ne $external.units) {
                return $external
            }
        }
        catch {}
    }
    return $defaults
}

function Normalize-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = ($Path -replace '/', '\').Trim()
    while ($p.Length -gt 3 -and $p.EndsWith('\')) {
        $p = $p.Substring(0, $p.Length - 1)
    }
    return $p
}

function Get-NormalizedPathForCompare {
    param([string]$Value)
    $tmp = Normalize-PathString $Value
    if ([string]::IsNullOrWhiteSpace($tmp)) { return '' }
    return $tmp.ToLowerInvariant()
}

function Test-PatternMatch {
    param([string]$Value, [object]$Patterns)
    if ([string]::IsNullOrWhiteSpace($Value) -or $null -eq $Patterns) { return $false }
    foreach ($pattern in @($Patterns)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$pattern) -and $Value -match [string]$pattern) {
            return $true
        }
    }
    return $false
}

function Get-DetectedUnitCode {
    param([string]$ComputerName, $ProfileConfig)

    $computerUpper = ([string]$ComputerName).ToUpperInvariant()
    if ($Unit -and $Unit.ToUpperInvariant() -ne 'AUTO') {
        return $Unit.ToUpperInvariant()
    }

    if ($null -ne $ProfileConfig -and $null -ne $ProfileConfig.units) {
        foreach ($property in $ProfileConfig.units.PSObject.Properties) {
            if (Test-PatternMatch -Value $computerUpper -Patterns $property.Value.computerPatterns) {
                return [string]$property.Name
            }
        }
    }

    if ($computerUpper -match '^ELCUN(\d+)-') {
        return ('UN' + $matches[1])
    }

    return 'UNKNOWN'
}

function Get-UnitProfile {
    param([string]$UnitCode, $ProfileConfig)
    if ($null -eq $ProfileConfig -or $null -eq $ProfileConfig.units) { return $null }
    if ($ProfileConfig.units.PSObject.Properties.Name -contains $UnitCode) {
        return $ProfileConfig.units.$UnitCode
    }
    return $null
}

function Get-RegistryValueSafe {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        return $item.$Name
    }
    catch {
        return $null
    }
}

function Get-BdeNetDirFromRegistry {
    foreach ($path in @(
        'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKCU:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT'
    )) {
        $value = Get-RegistryValueSafe -Path $path -Name 'NETDIR'
        if (-not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }
    return $null
}

function Resolve-ShareForLocalPath {
    param([string]$LocalPath, [array]$Shares, [string]$ComputerName)

    $normalized = Normalize-PathString $LocalPath
    if ([string]::IsNullOrWhiteSpace($normalized)) { return $null }

    if ($normalized -like '\\*') {
        return [pscustomobject]@{
            LocalPath = $normalized
            ShareName = $null
            SharePath = $null
            UncPath   = $normalized
        }
    }

    $best = $null
    foreach ($share in $Shares) {
        if (-not $share.Path) { continue }
        $sharePath = Normalize-PathString $share.Path
        if ([string]::IsNullOrWhiteSpace($sharePath)) { continue }
        if ($normalized.Length -lt $sharePath.Length) { continue }
        if (-not $normalized.StartsWith($sharePath, [System.StringComparison]::OrdinalIgnoreCase)) { continue }

        $boundaryOk = $false
        if ($normalized.Length -eq $sharePath.Length) {
            $boundaryOk = $true
        }
        elseif ($normalized.Substring($sharePath.Length, 1) -eq '\') {
            $boundaryOk = $true
        }

        if (-not $boundaryOk) { continue }

        if (($null -eq $best) -or ($sharePath.Length -gt $best.SharePath.Length)) {
            $best = [pscustomobject]@{
                ShareName = $share.Name
                SharePath = $sharePath
            }
        }
    }

    if ($null -eq $best) { return $null }

    $suffix = $normalized.Substring($best.SharePath.Length).TrimStart('\')
    $unc = "\\$ComputerName\$($best.ShareName)"
    if (-not [string]::IsNullOrWhiteSpace($suffix)) {
        $unc += "\$suffix"
    }

    return [pscustomobject]@{
        LocalPath = $normalized
        ShareName = $best.ShareName
        SharePath = $best.SharePath
        UncPath   = $unc
    }
}

function Resolve-Context {
    $profileConfig = Get-UnitProfiles
    $unitCode = Get-DetectedUnitCode -ComputerName $env:COMPUTERNAME -ProfileConfig $profileConfig
    $profile = Get-UnitProfile -UnitCode $unitCode -ProfileConfig $profileConfig

    $envUser = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', 'User')
    $envProcess = $env:HW_CAMINHO_DB
    $envMachine = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', 'Machine')

    $regDb = $null
    foreach ($path in @(
        'HKCU:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\Software\WOW6432Node\HeartWare\ECGV6\Geral'
    )) {
        $tmp = Get-RegistryValueSafe -Path $path -Name 'Caminho Database'
        if (-not [string]::IsNullOrWhiteSpace([string]$tmp)) {
            $regDb = [string]$tmp
            break
        }
    }

    $hwDb = if ($envProcess) { [string]$envProcess } elseif ($envUser) { [string]$envUser } elseif ($envMachine) { [string]$envMachine } elseif ($regDb) { [string]$regDb } elseif ($profile -and $profile.dbPath) { [string]$profile.dbPath } else { $null }
    $shares = @(Get-CimInstance -ClassName Win32_Share -ErrorAction SilentlyContinue | Where-Object { $_.Name -notmatch '\$$' -and $_.Name -ne 'IPC$' } | Select-Object Name, Path)

    $resolved = Resolve-ShareForLocalPath -LocalPath $hwDb -Shares $shares -ComputerName $env:COMPUTERNAME
    $effectiveRoot = if ($resolved -and $resolved.UncPath) { [string]$resolved.UncPath } elseif ($hwDb) { [string]$hwDb } elseif ($profile -and $profile.dbPath) { [string]$profile.dbPath } else { '' }
    $expectedNetDir = if ($profile -and $profile.netDirPath) { [string]$profile.netDirPath } elseif (-not [string]::IsNullOrWhiteSpace($effectiveRoot)) { (Join-Path $effectiveRoot 'NetDir') } else { '' }
    $currentNetDir = Get-BdeNetDirFromRegistry
    $bdeStatus = 'NAO_DETERMINADO'
    if ([string]::IsNullOrWhiteSpace($currentNetDir) -and -not [string]::IsNullOrWhiteSpace($expectedNetDir)) {
        $bdeStatus = 'AUSENTE'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($currentNetDir) -and -not [string]::IsNullOrWhiteSpace($expectedNetDir) -and (Get-NormalizedPathForCompare $currentNetDir) -eq (Get-NormalizedPathForCompare $expectedNetDir)) {
        $bdeStatus = 'OK'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($currentNetDir) -and -not [string]::IsNullOrWhiteSpace($expectedNetDir)) {
        $bdeStatus = 'DIVERGENTE'
    }

    return [pscustomobject]@{
        UnitCode = $unitCode
        Profile = $profile
        HwDb = $hwDb
        EffectiveRoot = $effectiveRoot
        ExpectedNetDir = $expectedNetDir
        CurrentNetDir = $currentNetDir
        BdeStatus = $bdeStatus
        NetDirExists = [bool](-not [string]::IsNullOrWhiteSpace($expectedNetDir) -and (Test-Path -LiteralPath $expectedNetDir))
        LockFileExists = [bool](-not [string]::IsNullOrWhiteSpace($expectedNetDir) -and (Test-Path -LiteralPath (Join-Path $expectedNetDir 'PDOXUSRS.NET')))
        IsAdmin = (Test-IsAdmin)
    }
}

Restart-WithBypassIfNeeded

if ($OpenRunbook -and (Test-Path -LiteralPath $RunbookPath)) {
    Start-Process -FilePath $RunbookPath | Out-Null
    exit 0
}

$needElevation = [bool]($Fix -or $CreateMissingDirs)
Restart-ElevatedIfNeeded -NeedElevation:$needElevation

if (-not (Test-Path -LiteralPath $OutputRoot)) {
    New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
}

$context = Resolve-Context
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$jsonOut = Join-Path $OutputRoot ("ECG_BDE_Fix_{0}_{1}.json" -f $env:COMPUTERNAME, $stamp)
$txtOut  = Join-Path $OutputRoot ("ECG_BDE_Fix_{0}_{1}.txt" -f $env:COMPUTERNAME, $stamp)

Write-Host "=== ECG BDE Fix ===" -ForegroundColor Cyan
Write-Host ""
Write-Host ("Unidade detectada  : {0}" -f $context.UnitCode)
Write-Host ("Perfil             : {0}" -f $(if ($context.Profile -and $context.Profile.name) { [string]$context.Profile.name } else { 'Sem perfil dedicado' }))
Write-Host ("HW_CAMINHO_DB      : {0}" -f $(if ($context.HwDb) { $context.HwDb } else { '<não definido>' }))
Write-Host ("Raiz efetiva       : {0}" -f $(if ($context.EffectiveRoot) { $context.EffectiveRoot } else { '<não resolvida>' }))
Write-Host ("NETDIR atual BDE   : {0}" -f $(if ($context.CurrentNetDir) { $context.CurrentNetDir } else { '<ausente>' }))
Write-Host ("NETDIR esperado    : {0}" -f $(if ($context.ExpectedNetDir) { $context.ExpectedNetDir } else { '<não determinado>' }))
Write-Host ("Status NETDIR      : {0}" -f $context.BdeStatus)
Write-Host ("NetDir existe      : {0}" -f $context.NetDirExists)
Write-Host ("PDOXUSRS.NET       : {0}" -f $context.LockFileExists)
Write-Host ""

$actions = New-Object System.Collections.Generic.List[string]
$backupFile = ''

if (-not [string]::IsNullOrWhiteSpace($context.ExpectedNetDir) -and ($context.BdeStatus -in @('AUSENTE','DIVERGENTE'))) {
    $actions.Add("Definir NETDIR no registro para $($context.ExpectedNetDir)")
}

if ($CreateMissingDirs -and -not [string]::IsNullOrWhiteSpace($context.ExpectedNetDir) -and -not $context.NetDirExists) {
    if ($context.ExpectedNetDir -like '\\*') {
        $actions.Add("Apenas registrar que o NetDir remoto está ausente. Criação automática remota foi bloqueada por segurança.")
    } else {
        $actions.Add("Criar diretório local do NetDir: $($context.ExpectedNetDir)")
    }
}

if ($CreateMissingDirs) {
    foreach ($localDir in @('C:\HW\NetDir', 'C:\HW\Private')) {
        if (-not (Test-Path -LiteralPath $localDir)) {
            $actions.Add("Criar diretório local padrão: $localDir")
        }
    }
}

if ($Fix) {
    if ($actions.Count -gt 0) {
        $backupDir = Join-Path $OutputRoot 'Backups'
        if (-not (Test-Path -LiteralPath $backupDir)) {
            New-Item -Path $backupDir -ItemType Directory -Force | Out-Null
        }
        $backupFile = Join-Path $backupDir ("BDE_INIT_{0}_{1}.reg" -f $env:COMPUTERNAME, $stamp)
        & reg.exe export "HKLM\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT" "$backupFile" /y | Out-Null

        if (-not [string]::IsNullOrWhiteSpace($context.ExpectedNetDir) -and ($context.BdeStatus -in @('AUSENTE','DIVERGENTE'))) {
            New-ItemProperty -Path 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT' -Name 'NETDIR' -Value $context.ExpectedNetDir -PropertyType String -Force | Out-Null
        }

        if ($CreateMissingDirs -and -not [string]::IsNullOrWhiteSpace($context.ExpectedNetDir) -and -not $context.NetDirExists -and ($context.ExpectedNetDir -notlike '\\*')) {
            New-Item -ItemType Directory -Path $context.ExpectedNetDir -Force | Out-Null
        }

        if ($CreateMissingDirs) {
            foreach ($localDir in @('C:\HW\NetDir', 'C:\HW\Private')) {
                if (-not (Test-Path -LiteralPath $localDir)) {
                    New-Item -ItemType Directory -Path $localDir -Force | Out-Null
                }
            }
        }

        $context = Resolve-Context
    }
}

$result = [pscustomobject]@{
    GeneratedAt = (Get-Date).ToString('o')
    ComputerName = $env:COMPUTERNAME
    UnitCode = $context.UnitCode
    ProfileName = $(if ($context.Profile -and $context.Profile.name) { [string]$context.Profile.name } else { 'Sem perfil dedicado' })
    HwDb = $context.HwDb
    EffectiveRoot = $context.EffectiveRoot
    ExpectedNetDir = $context.ExpectedNetDir
    CurrentNetDir = $context.CurrentNetDir
    BdeStatus = $context.BdeStatus
    NetDirExists = $context.NetDirExists
    LockFileExists = $context.LockFileExists
    FixMode = [bool]$Fix
    CreateMissingDirs = [bool]$CreateMissingDirs
    IsAdmin = [bool](Test-IsAdmin)
    PlannedOrAppliedActions = @($actions)
    BackupFile = $backupFile
    Recommendation = $(if ($Fix) { 'Reabrir o ECGV6 e, se necessário, reiniciar a estação.' } else { 'Usar o wrapper BAT/menu ou executar novamente com -Fix quando quiser aplicar a correção.' })
}

$result | ConvertTo-Json -Depth 8 | Set-Content -LiteralPath $jsonOut -Encoding UTF8

$txt = @()
$txt += '============================================'
$txt += 'ECG BDE FIX'
$txt += '============================================'
$txt += ('Computador          : {0}' -f $env:COMPUTERNAME)
$txt += ('Unidade             : {0}' -f $result.UnitCode)
$txt += ('Perfil              : {0}' -f $result.ProfileName)
$txt += ('HW_CAMINHO_DB       : {0}' -f $(if ($result.HwDb) { $result.HwDb } else { '<não definido>' }))
$txt += ('Raiz efetiva        : {0}' -f $(if ($result.EffectiveRoot) { $result.EffectiveRoot } else { '<não resolvida>' }))
$txt += ('NETDIR atual        : {0}' -f $(if ($result.CurrentNetDir) { $result.CurrentNetDir } else { '<ausente>' }))
$txt += ('NETDIR esperado     : {0}' -f $(if ($result.ExpectedNetDir) { $result.ExpectedNetDir } else { '<não determinado>' }))
$txt += ('Status              : {0}' -f $result.BdeStatus)
$txt += ('NetDir existe       : {0}' -f $result.NetDirExists)
$txt += ('PDOXUSRS.NET        : {0}' -f $result.LockFileExists)
$txt += ''
$txt += 'Ações'
$txt += '-----'
if ($actions.Count -gt 0) {
    foreach ($action in $actions) { $txt += ('- ' + $action) }
} else {
    $txt += '- Nenhuma ação necessária.'
}
if (-not [string]::IsNullOrWhiteSpace($backupFile)) {
    $txt += ''
    $txt += ('Backup de registro : {0}' -f $backupFile)
}
$txt += ''
$txt += ('Recomendação       : {0}' -f $result.Recommendation)
$txt += ('JSON               : {0}' -f $jsonOut)
$txt += ('TXT                : {0}' -f $txtOut)
$txt += '============================================'
$txt -join [Environment]::NewLine | Set-Content -LiteralPath $txtOut -Encoding UTF8

Write-Host ('JSON: ' + $jsonOut) -ForegroundColor Green
Write-Host ('TXT : ' + $txtOut) -ForegroundColor Green

if ($Fix) {
    Write-Host ''
    Write-Host 'Correção concluída. Reabra o ECGV6 e valide.' -ForegroundColor Yellow
} else {
    Write-Host ''
    Write-Host 'Modo diagnóstico concluído. Nenhuma alteração foi aplicada.' -ForegroundColor Yellow
}
