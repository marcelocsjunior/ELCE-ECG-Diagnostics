<#
.SYNOPSIS
  ECG profile builder v6.3.2 - target menu + custom + duration + source of truth no INI
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)][string]$TemplatePath,
    [Parameter(Mandatory=$true)][string]$OutputPath,
    [ValidateSet('Operational','Single','CompareBackend')][string]$Scenario = 'Operational',
    [string]$PrimaryTargetKey = 'WS2016',
    [string]$SelectedTargetsCsv = '',
    [int]$Minutes = 3,
    [int]$IntervalSeconds = 15,
    [string]$CustomLabel = 'CUSTOM_TARGET',
    [string]$CustomDbPath = '',
    [string]$CustomNetDirPath = '',
    [string]$CustomHostOsHint = 'Custom target',
    [string]$CustomSmbDialectHint = 'SMB custom',
    [switch]$CustomLegacyHint
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Read-IniFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "INI nao encontrado: $Path" }

    $ini = @{}
    $section = 'GeneralFlat'
    $ini[$section] = @{}

    foreach ($line in Get-Content -LiteralPath $Path -ErrorAction Stop) {
        $trimmed = [string]$line
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $trimmed = $trimmed.Trim()
        if ($trimmed.StartsWith('#') -or $trimmed.StartsWith(';')) { continue }

        if ($trimmed -match '^\[(.+)\]$') {
            $section = [string]$matches[1].Trim()
            if (-not $ini.ContainsKey($section)) { $ini[$section] = @{} }
            continue
        }

        if ($trimmed -match '^([^=]+?)=(.*)$') {
            $key = [string]$matches[1].Trim()
            $value = [string]$matches[2].Trim()
            if (-not $ini.ContainsKey($section)) { $ini[$section] = @{} }
            $ini[$section][$key] = $value
        }
    }

    return $ini
}

function Get-IniString {
    param($Ini, [string]$Section, [string]$Key, [string]$Default = '')
    if ($Ini.ContainsKey($Section) -and $Ini[$Section].ContainsKey($Key)) {
        $value = [string]$Ini[$Section][$Key]
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
    }
    return $Default
}

function Get-IniBool {
    param($Ini, [string]$Section, [string]$Key, [bool]$Default = $false)
    $raw = (Get-IniString -Ini $Ini -Section $Section -Key $Key -Default '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    if (@('1','true','yes','sim','y','on') -contains $raw) { return $true }
    if (@('0','false','no','nao','off') -contains $raw) { return $false }
    return $Default
}

function Normalize-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return '' }
    $p = ($Path -replace '/', '\').Trim()
    while ($p.Length -gt 3 -and $p.EndsWith('\')) { $p = $p.Substring(0, $p.Length - 1) }
    return $p
}

function Test-IsUncPath {
    param([string]$Path)
    $p = Normalize-PathString $Path
    return ($p -and $p.StartsWith('\\'))
}

function New-TargetObject {
    param(
        [string]$Key,
        [string]$Label,
        [string]$DbPath,
        [string]$NetDirPath,
        [string]$HostOsHint,
        [string]$SmbDialectHint,
        [bool]$LegacyHint,
        [string]$Notes,
        [string]$SectionName
    )

    return [PSCustomObject]@{
        Key = $Key
        Label = $Label
        DbPath = (Normalize-PathString $DbPath)
        NetDirPath = (Normalize-PathString $NetDirPath)
        HostOsHint = $HostOsHint
        SmbDialectHint = $SmbDialectHint
        LegacyHint = $LegacyHint
        Notes = $Notes
        SectionName = $SectionName
    }
}

function Get-PresetKeyFromSection {
    param($Ini, [string]$Section)
    $explicit = (Get-IniString -Ini $Ini -Section $Section -Key 'PresetKey' -Default '').ToUpperInvariant()
    if (-not [string]::IsNullOrWhiteSpace($explicit)) { return $explicit }

    switch ($Section) {
        'Target1' { return 'FS01' }
        'Target2' { return 'XP' }
        'Target3' { return 'WS2016' }
        'Target4' { return 'CUSTOM' }
    }

    $label = (Get-IniString -Ini $Ini -Section $Section -Key 'Label' -Default '').ToUpperInvariant()
    switch ($label) {
        'FS_FS01_LEGADO' { return 'FS01' }
        'FS_XP_LEGADO'   { return 'XP' }
        'FS_WS2016_BDE'  { return 'WS2016' }
        'CUSTOM_TARGET'  { return 'CUSTOM' }
    }

    return ''
}

function Get-TemplateTargets {
    param($Ini)

    $targets = @{}
    foreach ($section in @($Ini.Keys | Sort-Object)) {
        if ($section -notmatch '^Target\d+$') { continue }

        $key = Get-PresetKeyFromSection -Ini $Ini -Section $section
        if ([string]::IsNullOrWhiteSpace($key)) { continue }

        $targets[$key] = New-TargetObject `
            -Key $key `
            -Label (Get-IniString -Ini $Ini -Section $section -Key 'Label' -Default $section) `
            -DbPath (Get-IniString -Ini $Ini -Section $section -Key 'DbPath' -Default '') `
            -NetDirPath (Get-IniString -Ini $Ini -Section $section -Key 'NetDirPath' -Default '') `
            -HostOsHint (Get-IniString -Ini $Ini -Section $section -Key 'HostOsHint' -Default '') `
            -SmbDialectHint (Get-IniString -Ini $Ini -Section $section -Key 'SmbDialectHint' -Default '') `
            -LegacyHint (Get-IniBool -Ini $Ini -Section $section -Key 'LegacyHint' -Default $false) `
            -Notes (Get-IniString -Ini $Ini -Section $section -Key 'Notes' -Default '') `
            -SectionName $section
    }

    return $targets
}

function Add-UniqueTarget {
    param([System.Collections.Generic.List[object]]$List, $Target)
    foreach ($existing in $List) {
        if ([string]$existing.Key -eq [string]$Target.Key) { return }
    }
    [void]$List.Add($Target)
}

$ini = Read-IniFile -Path $TemplatePath
$presets = Get-TemplateTargets -Ini $ini

if ((Test-IsUncPath $CustomDbPath) -and (Test-IsUncPath $CustomNetDirPath)) {
    $presets['CUSTOM'] = New-TargetObject -Key 'CUSTOM' -Label $CustomLabel -DbPath $CustomDbPath -NetDirPath $CustomNetDirPath -HostOsHint $CustomHostOsHint -SmbDialectHint $CustomSmbDialectHint -LegacyHint ([bool]$CustomLegacyHint.IsPresent) -Notes 'Target custom informado no menu' -SectionName 'Target4'
}

$primaryKeyNormalized = ([string]$PrimaryTargetKey).Trim().ToUpperInvariant()
if ([string]::IsNullOrWhiteSpace($primaryKeyNormalized)) { $primaryKeyNormalized = 'WS2016' }

if ([string]::IsNullOrWhiteSpace($SelectedTargetsCsv)) {
    $selectedKeys = @($primaryKeyNormalized)
}
else {
    $selectedKeys = @(
        $SelectedTargetsCsv.Split(',') |
        ForEach-Object { ([string]$_).Trim().ToUpperInvariant() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
    )
}

$targets = New-Object System.Collections.Generic.List[object]
foreach ($key in $selectedKeys) {
    if ($presets.ContainsKey($key)) {
        Add-UniqueTarget -List $targets -Target $presets[$key]
    }
}

if ($targets.Count -eq 0) { throw 'Nenhum target valido selecionado.' }

if ($Scenario -eq 'CompareBackend' -and $targets.Count -lt 2) {
    if ($presets.ContainsKey($primaryKeyNormalized)) {
        Add-UniqueTarget -List $targets -Target $presets[$primaryKeyNormalized]
    }
}

if ($Scenario -eq 'CompareBackend' -and $targets.Count -lt 2) {
    throw 'CompareBackend requer pelo menos 2 targets validos ou 1 target + principal diferente.'
}

$primary = $targets[0]
foreach ($t in $targets) {
    if ([string]$t.Key -eq [string]$primaryKeyNormalized) {
        $primary = $t
        break
    }
}

$minutesValue = [Math]::Max(1, $Minutes)
$intervalValue = [Math]::Max(1, $IntervalSeconds)
$samplesValue = [Math]::Max(1, [int][Math]::Ceiling(($minutesValue * 60.0) / $intervalValue))
$selectedKeysEffective = @($targets | ForEach-Object { $_.Key }) -join ','

$lines = New-Object System.Collections.Generic.List[string]
[void]$lines.Add('# ECG Diagnostics Core v6.3.2 - perfil runtime gerado pelo menu')
[void]$lines.Add('# Detectar -> Decidir -> Relatar -> Corrigir somente se autorizada')
[void]$lines.Add('RuntimeProfileVersion=v6.3.2')
[void]$lines.Add('Scenario=' + $Scenario)
[void]$lines.Add('PrimaryTargetKey=' + $primary.Key)
[void]$lines.Add('SelectedTargetsCsv=' + $selectedKeysEffective)
[void]$lines.Add('ExpectedDbPath=' + $primary.DbPath)
[void]$lines.Add('ExpectedNetDir=' + $primary.NetDirPath)
[void]$lines.Add('ExpectedExePath=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe'))
[void]$lines.Add('SetMachineHwPath=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'SetMachineHwPath' -Default 'true'))
[void]$lines.Add('OutDir=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'OutDir' -Default 'C:\ECG\FieldKit\out'))
[void]$lines.Add('StationRole=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'StationRole' -Default 'AUTO'))
[void]$lines.Add('StationAlias=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'StationAlias' -Default 'VIEWER'))
[void]$lines.Add('MonitorMinutes=' + [string]$minutesValue)
[void]$lines.Add('SampleIntervalSeconds=' + [string]$intervalValue)
[void]$lines.Add('CpuProcessCaptureThreshold=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'CpuProcessCaptureThreshold' -Default '80'))
[void]$lines.Add('TopProcessCaptureCount=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'TopProcessCaptureCount' -Default '3'))
[void]$lines.Add('EnableLatencyMetrics=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'EnableLatencyMetrics' -Default 'true'))
[void]$lines.Add('EnableEcgProcessMetrics=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'EnableEcgProcessMetrics' -Default 'true'))
[void]$lines.Add('EnableDiskMetrics=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'EnableDiskMetrics' -Default 'false'))
[void]$lines.Add('EnableNetworkMetrics=' + (Get-IniString -Ini $ini -Section 'GeneralFlat' -Key 'EnableNetworkMetrics' -Default 'false'))
[void]$lines.Add('')
[void]$lines.Add('[General]')
[void]$lines.Add('OutDir=' + (Get-IniString -Ini $ini -Section 'General' -Key 'OutDir' -Default 'C:\ECG\FieldKit\out'))
[void]$lines.Add('LocalProbePath=' + (Get-IniString -Ini $ini -Section 'General' -Key 'LocalProbePath' -Default 'C:\Windows'))
[void]$lines.Add('Samples=' + [string]$samplesValue)
[void]$lines.Add('IntervalSeconds=' + [string]$intervalValue)
[void]$lines.Add('PingCount=' + (Get-IniString -Ini $ini -Section 'General' -Key 'PingCount' -Default '4'))
[void]$lines.Add('TcpTimeoutMs=' + (Get-IniString -Ini $ini -Section 'General' -Key 'TcpTimeoutMs' -Default '1500'))
[void]$lines.Add('WorkloadLabel=' + (Get-IniString -Ini $ini -Section 'General' -Key 'WorkloadLabel' -Default 'ECGv6 DBE/BDE'))
[void]$lines.Add('ExpectedExePath=' + (Get-IniString -Ini $ini -Section 'General' -Key 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe'))

$index = 1
foreach ($target in $targets) {
    [void]$lines.Add('')
    [void]$lines.Add('[Target' + [string]$index + ']')
    [void]$lines.Add('Enabled=true')
    [void]$lines.Add('PresetKey=' + $target.Key)
    [void]$lines.Add('Label=' + $target.Label)
    [void]$lines.Add('DbPath=' + $target.DbPath)
    [void]$lines.Add('NetDirPath=' + $target.NetDirPath)
    [void]$lines.Add('HostOsHint=' + $target.HostOsHint)
    [void]$lines.Add('SmbDialectHint=' + $target.SmbDialectHint)
    [void]$lines.Add('LegacyHint=' + ($(if ($target.LegacyHint) { 'true' } else { 'false' })))
    [void]$lines.Add('Notes=' + $target.Notes)
    $index++
}

$dir = Split-Path -Path $OutputPath -Parent
if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
    New-Item -ItemType Directory -Path $dir -Force | Out-Null
}

[System.IO.File]::WriteAllText($OutputPath, ($lines -join [Environment]::NewLine), (New-Object System.Text.UTF8Encoding($false)))
Write-Host $OutputPath
