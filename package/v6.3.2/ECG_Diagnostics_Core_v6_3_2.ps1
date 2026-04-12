<#
.SYNOPSIS
    ECG Diagnostics Core - suite unificada para diagnostico, correcao e comparativo de backend ECG/BDE.
.DESCRIPTION
    Modos:
      Fix          : corrige NETDIR (HKLM/HKCU), IDAPI32.CFG + timeline + relatorio HTML/JSON
      Auto         : somente diagnostico (sem correcoes) + timeline + relatorio HTML/JSON
      Compare      : compara dois laudos JSON (modo legado ECG_Report.json)
      Single       : diagnostico single-target do primeiro backend habilitado ou do TargetLabel informado
      CompareBackend: compara todos os backends habilitados no mesmo run
      CompareJson  : compara dois relatorios JSON gerados pela propria suite
      Rollback     : restaura backup .reg do BDE
      Monitor      : monitoramento prolongado com grafico, hipoteses e score
      CollectStatic: coleta estatica de informacoes (JSON)
      Hotfix      : remove regressao de parse por duplo BOM UTF-8 no topo do arquivo
.NOTES
    Versao: 6.3.2-hardening
#>

[CmdletBinding()]
param(
    [ValidateSet('Fix','Auto','Detect','Compare','CompareJson','Rollback','Monitor','CollectStatic','Single','CompareBackend')]
    [string]$Mode = 'Auto',

    [string]$ProfilePath = '',
    [string]$OutDir = 'C:\ECG\FieldKit\out',
    [string]$RollbackFile = '',
    [string]$CompareLeftReport = '',
    [string]$CompareRightReport = '',
    [string]$TargetLabel = '',
    [switch]$OpenReport,
    [switch]$AuthorizedRemediation
)

$ErrorActionPreference = 'Stop'
$script:ToolName = 'ECG Diagnostics Core'
$script:ToolVersion = '6.3.2-hardening'
$script:ToolVersionShort = 'v6.3.2'
$script:HostName = $env:COMPUTERNAME
$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:ScriptDir = Split-Path -Parent $script:ScriptPath
$script:LogLines = New-Object System.Collections.ArrayList
$script:RunId = (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:HostName
$script:CurrentRunRoot = ''
$script:CliOutDirProvided = $PSBoundParameters.ContainsKey('OutDir')

if ([string]::IsNullOrWhiteSpace($ProfilePath)) {
    try {
        $candidateProfiles = @(
            (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_3_2.ini'),
            (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_3_1.ini'),
            (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_2.ini'),
            (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_1.ini'),
            (Join-Path $script:ScriptDir 'ECG_FieldKit_Compare_v6.ini'),
            (Join-Path $script:ScriptDir 'ECG_FieldKit.ini')
        )
        foreach ($candidateProfile in $candidateProfiles) {
            if (Test-Path -LiteralPath $candidateProfile) {
                $ProfilePath = $candidateProfile
                break
            }
        }
    }
    catch {}
}

function Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [void]$script:LogLines.Add($line)
    Write-Host $line
}

function Write-Utf8File {
    param([string]$Path, [string]$Text)
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Write-UnicodeFile {
    param([string]$Path, [string]$Text)
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $enc = New-Object System.Text.UnicodeEncoding($false, $true)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
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

function Get-RegistryValueSafe {
    param([string]$Path, [string]$Name)
    try {
        $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
        return $item.$Name
    }
    catch {
        return $null
    }
}

function Set-RegistryValueSafe {
    param([string]$Path, [string]$Name, [string]$Value)
    try {
        if (-not (Test-Path -LiteralPath $Path)) {
            New-Item -Path $Path -Force | Out-Null
        }
        $existing = Get-RegistryValueSafe -Path $Path -Name $Name
        if ($null -ne $existing) {
            Set-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -ErrorAction Stop
        }
        else {
            New-ItemProperty -LiteralPath $Path -Name $Name -Value $Value -PropertyType String -Force -ErrorAction Stop | Out-Null
        }
        return $true
    }
    catch {
        return $false
    }
}

function Read-ProfileIni {
    param([string]$Path)
    $profile = @{}
    if ($Path -and (Test-Path $Path)) {
        foreach ($line in (Get-Content -Path $Path -ErrorAction Stop)) {
            if ($line -match '^\s*([^#=;]+)\s*=\s*(.*)\s*$') {
                $profile[$matches[1].Trim()] = $matches[2].Trim()
            }
        }
    }
    return $profile
}

function Get-ProfileString {
    param($Profile, [string]$Name, [string]$Default = '')
    if ($Profile -and $Profile.ContainsKey($Name)) {
        $value = [string]$Profile[$Name]
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
    }
    return $Default
}

function Get-ProfileInt {
    param($Profile, [string]$Name, [int]$Default)
    $raw = Get-ProfileString -Profile $Profile -Name $Name -Default ''
    $tmp = 0
    if ([int]::TryParse($raw, [ref]$tmp)) { return $tmp }
    return $Default
}

function Get-ProfileBool {
    param($Profile, [string]$Name, [bool]$Default = $false)
    $raw = (Get-ProfileString -Profile $Profile -Name $Name -Default '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    if (@('1','true','yes','sim','y','on') -contains $raw) { return $true }
    if (@('0','false','no','nao','nao','off') -contains $raw) { return $false }
    return $Default
}

function Get-MachineTypeFromProfileOrName {
    param(
        [string]$ComputerName,
        $Profile
    )

    $computerUpper = ([string]$ComputerName).ToUpperInvariant()
    $stationRole = (Get-ProfileString -Profile $Profile -Name 'StationRole' -Default 'AUTO').ToUpperInvariant()
    $stationAlias = (Get-ProfileString -Profile $Profile -Name 'StationAlias' -Default '').ToUpperInvariant()

    switch ($stationRole) {
        'EXECUTANTE' { return 'Estacao de exames' }
        'VIEWER'     { return 'Estacao de visualizacao' }
        'HOST_XP'    { return 'Host XP legado' }
    }

    if ($computerUpper -match 'SRVVM1-FS01') { return 'Servidor de arquivos' }
    if ($computerUpper -match '(^|[-_])ECG([0-9A-Z_-]*$)') { return 'Estacao de exames' }
    if ($computerUpper -match '(^|[-_])(CST|CON|VIEW)([0-9A-Z_-]*$)') { return 'Estacao de visualizacao' }

    switch ($stationAlias) {
        'EXECUTANTE' { return 'Estacao de exames' }
        'VIEWER'     { return 'Estacao de visualizacao' }
        'HOST_XP'    { return 'Host XP legado' }
    }

    return 'Estacao de trabalho'
}

function Resolve-EffectiveOutDir {
    param($Profile)
    if ($script:CliOutDirProvided -and -not [string]::IsNullOrWhiteSpace($OutDir)) { return $OutDir }
    $profileOutDir = Get-ProfileString -Profile $Profile -Name 'OutDir' -Default ''
    if (-not [string]::IsNullOrWhiteSpace($profileOutDir)) { return $profileOutDir }
    return $OutDir
}

function Normalize-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = ($Path -replace '/', '\').Trim()
    while ($p.Length -gt 3 -and $p.EndsWith('\')) { $p = $p.Substring(0, $p.Length - 1) }
    return $p
}

function Test-IsUncPath {
    param([string]$Path)
    $p = Normalize-PathString $Path
    return ($p -and $p.StartsWith('\\'))
}

function ConvertTo-RegEscapedString {
    param([string]$Value)
    if ($null -eq $Value) { return $null }
    return ($Value -replace '\\', '\\\\' -replace '"', '\\"')
}

function Export-BdeRollbackFile {
    param([string]$Path)
    $entries = @(
        @{ PsPath = 'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT'; RegPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT' },
        @{ PsPath = 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT'; RegPath = 'HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT' },
        @{ PsPath = 'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT'; RegPath = 'HKEY_CURRENT_USER\Software\Borland\Database Engine\Settings\SYSTEM\INIT' }
    )
    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add('Windows Registry Editor Version 5.00')
    [void]$lines.Add('')
    foreach ($entry in $entries) {
        [void]$lines.Add("[$($entry.RegPath)]")
        $current = Get-RegistryValueSafe -Path $entry.PsPath -Name 'NETDIR'
        if ($null -eq $current -or [string]::IsNullOrWhiteSpace([string]$current)) {
            [void]$lines.Add('"NETDIR"=-')
        }
        else {
            $escaped = ConvertTo-RegEscapedString -Value ([string]$current)
            [void]$lines.Add('"NETDIR"="' + $escaped + '"')
        }
        [void]$lines.Add('')
    }
    Write-UnicodeFile -Path $Path -Text ($lines -join [Environment]::NewLine)
    return $Path
}

function Set-HwCaminhoDbSafe {
    param([string]$Value)
    $changes = @()
    try {
        $env:HW_CAMINHO_DB = $Value
        $changes += 'HW_CAMINHO_DB ajustado no processo atual'
    }
    catch {}
    try {
        [Environment]::SetEnvironmentVariable('HW_CAMINHO_DB', $Value, 'User')
        $changes += 'HW_CAMINHO_DB ajustado para o usuario atual'
    }
    catch {}
    if (Test-IsAdmin) {
        try {
            [Environment]::SetEnvironmentVariable('HW_CAMINHO_DB', $Value, 'Machine')
            $changes += 'HW_CAMINHO_DB ajustado em nivel de maquina'
        }
        catch {}
    }
    return ,$changes
}

function Get-WmiSafe {
    param([string]$ClassName)
    try { return @(Get-WmiObject -Class $ClassName -ErrorAction Stop) } catch { return @() }
}

function Get-LocalShares {
    $rows = @(
        Get-WmiSafe -ClassName 'Win32_Share' |
        Select-Object -Property @('Name','Path')
    )
    return $rows
}

function Get-NetworkConnections {
    $rows = @(Get-WmiSafe -ClassName 'Win32_NetworkConnection')
    return $rows
}

function Test-CommandExists {
    param([string]$Name)
    try { return $null -ne (Get-Command -Name $Name -ErrorAction Stop) } catch { return $false }
}

function ConvertTo-DoubleSafe {
    param([object]$Value)
    if ($null -eq $Value) { return $null }
    try { return [double]$Value } catch { return $null }
}

function Get-ArrayStats {
    param([object[]]$Values)
    $nums = New-Object System.Collections.Generic.List[double]
    foreach ($v in @($Values)) {
        $d = ConvertTo-DoubleSafe -Value $v
        if ($null -ne $d) { [void]$nums.Add([double]$d) }
    }
    if ($nums.Count -eq 0) {
        return [PSCustomObject]@{ Count = 0; Average = $null; Maximum = $null; Minimum = $null }
    }
    $sum = 0.0
    $max = $nums[0]
    $min = $nums[0]
    foreach ($n in $nums) {
        $sum += $n
        if ($n -gt $max) { $max = $n }
        if ($n -lt $min) { $min = $n }
    }
    return [PSCustomObject]@{
        Count = $nums.Count
        Average = [math]::Round(($sum / $nums.Count), 2)
        Maximum = [math]::Round($max, 2)
        Minimum = [math]::Round($min, 2)
    }
}

function Get-UncHostFromPath {
    param([string]$Path)
    $p = Normalize-PathString $Path
    if ($p -and $p -match '^\\\\([^\\]+)\\') { return $matches[1] }
    return $null
}

function Get-CpuPercentSafe {
    try {
        $sample = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples[0].CookedValue
        return [math]::Round([double]$sample, 2)
    }
    catch {
        try {
            if (Test-CommandExists -Name 'Get-CimInstance') {
                $cpu = (Get-CimInstance Win32_Processor -ErrorAction Stop | ForEach-Object { $_.LoadPercentage } | Measure-Object -Average).Average
                if ($null -ne $cpu) { return [math]::Round([double]$cpu, 2) }
            }
        }
        catch {}
        try {
            $cpu2 = (Get-WmiObject Win32_Processor -ErrorAction Stop | ForEach-Object { $_.LoadPercentage } | Measure-Object -Average).Average
            if ($null -ne $cpu2) { return [math]::Round([double]$cpu2, 2) }
        }
        catch {}
        return $null
    }
}


function Get-LogicalProcessorCountSafe {
    try {
        if (Test-CommandExists -Name 'Get-CimInstance') {
            $cs = Get-CimInstance Win32_ComputerSystem -ErrorAction Stop
            if ($cs -and $cs.NumberOfLogicalProcessors -and [int]$cs.NumberOfLogicalProcessors -gt 0) { return [int]$cs.NumberOfLogicalProcessors }
        }
    }
    catch {}
    try {
        $cs2 = Get-WmiObject Win32_ComputerSystem -ErrorAction Stop
        if ($cs2 -and $cs2.NumberOfLogicalProcessors -and [int]$cs2.NumberOfLogicalProcessors -gt 0) { return [int]$cs2.NumberOfLogicalProcessors }
    }
    catch {}
    return 1
}

function Measure-PathProbeSafe {
    param([string]$Path)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $ok = $false
    try {
        $ok = Test-Path $Path
    }
    catch {
        $ok = $false
    }
    $sw.Stop()
    return [PSCustomObject]@{
        Accessible = $ok
        ProbeMs = [math]::Round([double]$sw.Elapsed.TotalMilliseconds, 2)
    }
}

function Get-PercentileSafe {
    param([object[]]$Values, [double]$Percentile = 95)
    $nums = @()
    foreach ($v in @($Values)) {
        $d = ConvertTo-DoubleSafe -Value $v
        if ($null -ne $d) { $nums += [double]$d }
    }
    if ($nums.Count -eq 0) { return $null }
    $ordered = @($nums | Sort-Object)
    if ($ordered.Count -eq 1) { return [math]::Round([double]$ordered[0], 2) }
    $p = [Math]::Max(0, [Math]::Min(100, [double]$Percentile))
    $index = [math]::Ceiling(($p / 100) * $ordered.Count) - 1
    if ($index -lt 0) { $index = 0 }
    if ($index -ge $ordered.Count) { $index = $ordered.Count - 1 }
    return [math]::Round([double]$ordered[$index], 2)
}

function Get-PerformanceCounterSamplesSafe {
    param([string[]]$CounterPaths)
    if (-not (Test-CommandExists -Name 'Get-Counter')) { return @() }
    try {
        return @((Get-Counter -Counter $CounterPaths -ErrorAction Stop).CounterSamples)
    }
    catch {
        return @()
    }
}

function Get-PhysicalDiskQueueLengthSafe {
    $samples = @(Get-PerformanceCounterSamplesSafe -CounterPaths @('\PhysicalDisk(_Total)\Avg. Disk Queue Length'))
    if ($samples.Count -eq 0) { return $null }
    $value = ConvertTo-DoubleSafe -Value $samples[0].CookedValue
    if ($null -eq $value) { return $null }
    return [math]::Round([double]$value, 2)
}

function Get-NetworkBytesTotalPerSecSafe {
    $samples = @(Get-PerformanceCounterSamplesSafe -CounterPaths @('\Network Interface(*)\Bytes Total/sec'))
    if ($samples.Count -eq 0) { return $null }
    $sum = 0.0
    $found = $false
    foreach ($sample in $samples) {
        $path = [string]$sample.Path
        if ($path -match '(?i)loopback|isatap|teredo|pseudo|tunnel') { continue }
        $value = ConvertTo-DoubleSafe -Value $sample.CookedValue
        if ($null -ne $value) {
            $sum += [double]$value
            $found = $true
        }
    }
    if (-not $found) { return $null }
    return [math]::Round($sum, 2)
}

function Get-ProcessCounterCpuPercentSafe {
    param([string]$ProcessPattern = 'ECGV6*')
    if (-not (Test-CommandExists -Name 'Get-Counter')) { return $null }
    $counterPath = ('\Process({0})\% Processor Time' -f $ProcessPattern)
    try {
        $samples = @((Get-Counter -Counter $counterPath -ErrorAction Stop).CounterSamples)
        if ($samples.Count -eq 0) { return $null }
        $logicalProcessors = Get-LogicalProcessorCountSafe
        $sum = 0.0
        $found = $false
        foreach ($sample in $samples) {
            $path = [string]$sample.Path
            if ($path -match '(?i)_Total|Idle') { continue }
            $value = ConvertTo-DoubleSafe -Value $sample.CookedValue
            if ($null -ne $value) {
                $sum += [double]$value
                $found = $true
            }
        }
        if (-not $found) { return $null }
        if ($logicalProcessors -gt 0) { $sum = $sum / $logicalProcessors }
        if ($sum -lt 0) { $sum = 0 }
        return [math]::Round($sum, 2)
    }
    catch {
        return $null
    }
}

function Get-ProcessOwnerSafe {
    param([int]$ProcessId)
    try {
        $proc = Get-WmiObject Win32_Process -Filter "ProcessId = $ProcessId" -ErrorAction Stop
        if ($null -eq $proc) { return $null }
        $owner = $proc.GetOwner()
        if ($owner -and $owner.ReturnValue -eq 0) {
            $domain = [string]$owner.Domain
            $user = [string]$owner.User
            if (-not [string]::IsNullOrWhiteSpace($domain) -and -not [string]::IsNullOrWhiteSpace($user)) { return "$domain\$user" }
            if (-not [string]::IsNullOrWhiteSpace($user)) { return $user }
        }
    }
    catch {}
    return $null
}

function Test-IsEcgRelatedProcess {
    param([string]$ProcessName, [string]$ExecutablePath = '')
    $name = [string]$ProcessName
    $path = [string]$ExecutablePath
    if ($name -match '(?i)(^|[^a-z])(ecg|ecgv6|idapi|bde|pdox|hw)([^a-z]|$)') { return $true }
    if ($path -match '(?i)\\HW\\ECG\\|ECGV6\.exe|IDAPI|Borland Shared\\BDE|Paradox') { return $true }
    return $false
}

function Get-TopCpuProcessesSafe {
    param([int]$Top = 3)

    $limit = [Math]::Min(5, [Math]::Max(1, $Top))
    $results = @()
    $logicalProcessors = Get-LogicalProcessorCountSafe

    try {
        $perfRows = @(Get-WmiObject -Class Win32_PerfFormattedData_PerfProc_Process -ErrorAction Stop |
            Where-Object {
                $_.Name -and
                $_.Name -notin @('_Total','Idle') -and
                $null -ne $_.IDProcess -and
                [int]$_.IDProcess -gt 0
            } |
            Sort-Object -Property PercentProcessorTime -Descending |
            Select-Object -First $limit)

        foreach ($row in $perfRows) {
            $pid = 0
            try { $pid = [int]$row.IDProcess } catch { $pid = 0 }
            if ($pid -le 0) { continue }

            $owner = $null
            $path = $null
            try {
                $procWmi = Get-WmiObject Win32_Process -Filter "ProcessId = $pid" -ErrorAction Stop
                if ($procWmi) {
                    $path = $procWmi.ExecutablePath
                    $ownerInfo = $procWmi.GetOwner()
                    if ($ownerInfo -and $ownerInfo.ReturnValue -eq 0) {
                        if (-not [string]::IsNullOrWhiteSpace([string]$ownerInfo.Domain) -and -not [string]::IsNullOrWhiteSpace([string]$ownerInfo.User)) {
                            $owner = "$($ownerInfo.Domain)\$($ownerInfo.User)"
                        }
                        elseif (-not [string]::IsNullOrWhiteSpace([string]$ownerInfo.User)) {
                            $owner = [string]$ownerInfo.User
                        }
                    }
                }
            }
            catch {}

            if ([string]::IsNullOrWhiteSpace([string]$owner)) {
                $owner = Get-ProcessOwnerSafe -ProcessId $pid
            }

            $cpuPercent = $null
            try {
                $rawCpu = [double]$row.PercentProcessorTime
                if ($logicalProcessors -gt 0) { $rawCpu = $rawCpu / $logicalProcessors }
                $cpuPercent = [math]::Round($rawCpu, 2)
            }
            catch {}

            $results += [PSCustomObject]@{
                Name = [string]$row.Name
                ProcessId = $pid
                CpuPercent = $cpuPercent
                Owner = $owner
                Path = $path
                IsEcgRelated = (Test-IsEcgRelatedProcess -ProcessName ([string]$row.Name) -ExecutablePath ([string]$path))
            }
        }
    }
    catch {}

    if (@($results).Count -gt 0) {
        return @($results | Sort-Object -Property @{ Expression = { if ($null -ne $_.CpuPercent) { [double]$_.CpuPercent } else { -1 } }; Descending = $true } | Select-Object -First $limit)
    }

    if (Test-CommandExists -Name 'Get-Counter') {
        try {
            $samples = @((Get-Counter -Counter @('\Process(*)\ID Process','\Process(*)\% Processor Time') -ErrorAction Stop).CounterSamples)
            if ($samples.Count -gt 0) {
                $map = @{}
                foreach ($sample in $samples) {
                    $path = [string]$sample.Path
                    if ($path -notmatch '\\Process\((.+?)\)\\(.+)$') { continue }
                    $instanceName = [string]$matches[1]
                    if ($instanceName -match '^(Idle|_Total)$') { continue }
                    if (-not $map.ContainsKey($instanceName)) {
                        $map[$instanceName] = [ordered]@{ Name = $instanceName; ProcessId = $null; CpuPercent = $null }
                    }
                    if ($path -like '*\ID Process') {
                        try { $map[$instanceName].ProcessId = [int]$sample.CookedValue } catch {}
                    }
                    elseif ($path -like '*\% Processor Time') {
                        try {
                            $rawCpu = [double]$sample.CookedValue
                            if ($logicalProcessors -gt 0) { $rawCpu = $rawCpu / $logicalProcessors }
                            $map[$instanceName].CpuPercent = [math]::Round($rawCpu, 2)
                        }
                        catch {}
                    }
                }

                $fallback = @()
                foreach ($entry in $map.GetEnumerator()) {
                    $procName = [string]$entry.Value.Name
                    $pid = 0
                    try { $pid = [int]$entry.Value.ProcessId } catch { $pid = 0 }
                    if ($pid -le 0) { continue }
                    $owner = Get-ProcessOwnerSafe -ProcessId $pid
                    $path = $null
                    try {
                        $procWmi = Get-WmiObject Win32_Process -Filter "ProcessId = $pid" -ErrorAction Stop
                        if ($procWmi) { $path = $procWmi.ExecutablePath }
                    }
                    catch {}
                    $fallback += [PSCustomObject]@{
                        Name = $procName
                        ProcessId = $pid
                        CpuPercent = $entry.Value.CpuPercent
                        Owner = $owner
                        Path = $path
                        IsEcgRelated = (Test-IsEcgRelatedProcess -ProcessName $procName -ExecutablePath ([string]$path))
                    }
                }
                if ($fallback.Count -gt 0) {
                    return @($fallback | Sort-Object -Property @{ Expression = { if ($null -ne $_.CpuPercent) { [double]$_.CpuPercent } else { -1 } }; Descending = $true } | Select-Object -First $limit)
                }
            }
        }
        catch {}
    }

    return @()
}

function Get-CpuResponsibleProcessSummary {
    param($ProcessInfo, [switch]$IncludeOwner = $true, [switch]$IncludeCpu = $true)
    if ($null -eq $ProcessInfo) { return $null }

    $parts = New-Object System.Collections.Generic.List[string]
    $name = [string]$ProcessInfo.Name
    if (-not [string]::IsNullOrWhiteSpace($name)) { [void]$parts.Add($name) }
    if ($null -ne $ProcessInfo.ProcessId) { [void]$parts.Add("PID $($ProcessInfo.ProcessId)") }
    if ($IncludeCpu -and $null -ne $ProcessInfo.CpuPercent) { [void]$parts.Add("CPU proc $($ProcessInfo.CpuPercent)%") }
    if ($IncludeOwner -and -not [string]::IsNullOrWhiteSpace([string]$ProcessInfo.Owner)) { [void]$parts.Add([string]$ProcessInfo.Owner) }
    if ($ProcessInfo.IsEcgRelated) { [void]$parts.Add('relacionado ao ECG') }
    if ($parts.Count -eq 0) { return $null }
    return ($parts -join ' | ')
}

function Get-SmbMetricsSafe {
    param($MachineType, [string]$DbHost = '', [string]$NetDirHost = '')
    $result = [ordered]@{
        SmbConnectionCount = $null
        SmbSessionCount = $null
        RelevantOpenFileCount = $null
        RelevantNetDirOpenFileCount = $null
        SmbQueryTimedOut = $false
    }

    $hosts = @()
    foreach ($h in @($DbHost, $NetDirHost, 'SRVVM1-FS01', '192.168.1.57')) {
        if (-not [string]::IsNullOrWhiteSpace($h) -and ($hosts -notcontains $h)) { $hosts += $h }
    }
    $normalizedHosts = @($hosts | ForEach-Object { ([string]$_).ToUpperInvariant() })

    try {
        if ($MachineType -eq 'Servidor de arquivos') {
            if (Test-CommandExists -Name 'Get-SmbSession') {
                $sessions = @(Get-SmbSession -ErrorAction SilentlyContinue)
                $result.SmbSessionCount = $sessions.Count
            }
            if (Test-CommandExists -Name 'Get-SmbOpenFile') {
                $openFiles = @(Get-SmbOpenFile -ErrorAction SilentlyContinue)
                $result.RelevantOpenFileCount = @($openFiles | Where-Object { $_.Path -like '*\Database*' }).Count
                $result.RelevantNetDirOpenFileCount = @($openFiles | Where-Object { $_.Path -like '*\NetDir*' -or $_.Path -like '*\Netdir*' }).Count
            }
        }
        else {
            if (Test-CommandExists -Name 'Get-SmbConnection') {
                $allConnections = @(Get-SmbConnection -ErrorAction SilentlyContinue)
                if ($normalizedHosts.Count -gt 0) {
                    $connections = @($allConnections | Where-Object {
                        $serverName = ''
                        if ($null -ne $_.ServerName) { $serverName = ([string]$_.ServerName).ToUpperInvariant() }
                        $serverShort = if ($serverName -match '^([^\.]+)\.') { $matches[1] } else { $serverName }
                        ($normalizedHosts -contains $serverName) -or ($normalizedHosts -contains $serverShort)
                    })
                }
                else {
                    $connections = $allConnections
                }
                $result.SmbConnectionCount = $connections.Count
            }
        }
    }
    catch {
        $result.SmbQueryTimedOut = $true
    }

    return [PSCustomObject]$result
}

function Get-LockFileCountSafe {
    param([string]$NetDirPath)
    try {
        if (Test-Path $NetDirPath) {
            return @(Get-ChildItem -Path $NetDirPath -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -match 'PDOXUSRS|\.LCK$|\.NET$' }).Count
        }
    }
    catch {}
    return $null
}

function Get-KnownMachineInfo {
    param(
        [string]$ComputerName,
        $Profile = $null
    )
    $machineType = Get-MachineTypeFromProfileOrName -ComputerName $ComputerName -Profile $Profile
    $executedBy = try { [Security.Principal.WindowsIdentity]::GetCurrent().Name } catch { $env:USERNAME }
    return [PSCustomObject]@{
        ComputerName = $ComputerName
        MachineType = $machineType
        ExecutedBy = $executedBy
    }
}


function Get-IdapiCfgCandidatePaths {
    return @(
        'C:\Program Files (x86)\Common Files\Borland Shared\BDE\IDAPI32.CFG',
        'C:\Program Files\Common Files\Borland Shared\BDE\IDAPI32.CFG'
    )
}

function Get-IdapiCfgNetDirValue {
    param([string]$CfgPath)
    if ([string]::IsNullOrWhiteSpace($CfgPath) -or -not (Test-Path -LiteralPath $CfgPath)) { return $null }
    try {
        $content = [System.IO.File]::ReadAllText($CfgPath, [System.Text.Encoding]::Default)
        foreach ($pattern in @('(?im)^\s*NET\s+DIR\s*=\s*(.+)$','(?im)^\s*NETDIR\s*=\s*(.+)$')) {
            $match = [regex]::Match($content, $pattern)
            if ($match.Success) {
                return Normalize-PathString ($match.Groups[1].Value.Trim())
            }
        }
    }
    catch {}
    return ''
}

function Get-BdeSourceSnapshot {
    param($Profile)
    $officialDb = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedDbPath' -Default '\\192.168.1.57\Database')
    $officialNetDir = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedNetDir' -Default (Join-Path $officialDb 'NetDir'))
    $officialExe = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe')
    $hwDb = Normalize-PathString $env:HW_CAMINHO_DB
    $hwNet = if ($hwDb) { Normalize-PathString (Join-Path $hwDb 'NetDir') } else { '' }
    $hklm = 'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT'
    $hklmWow = 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT'
    $hkcu = 'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT'
    $cfgEntries = @()
    foreach ($cfgPath in @(Get-IdapiCfgCandidatePaths)) {
        if (Test-Path -LiteralPath $cfgPath) {
            $cfgNetDir = Normalize-PathString (Get-IdapiCfgNetDirValue -CfgPath $cfgPath)
            $cfgEntries += [PSCustomObject]@{
                Path = $cfgPath
                NetDir = $cfgNetDir
                Accessible = $(if ($cfgNetDir) { Test-Path -LiteralPath $cfgNetDir } else { $false })
            }
        }
    }
    return [PSCustomObject]@{
        OfficialDbPath = $officialDb
        OfficialNetDir = $officialNetDir
        OfficialExePath = $officialExe
        DesiredDbAccessible = (Test-Path -LiteralPath $officialDb)
        DesiredNetDirAccessible = (Test-Path -LiteralPath $officialNetDir)
        DesiredExeAccessible = (Test-Path -LiteralPath $officialExe)
        HwDb = $hwDb
        HwDbAccessible = $(if ($hwDb) { Test-Path -LiteralPath $hwDb } else { $false })
        HwNetDir = $hwNet
        HwNetDirAccessible = $(if ($hwNet) { Test-Path -LiteralPath $hwNet } else { $false })
        HkcuPath = $hkcu
        HkcuNetDir = Normalize-PathString (Get-RegistryValueSafe -Path $hkcu -Name 'NETDIR')
        HkcuNetDirAccessible = $(if (Get-RegistryValueSafe -Path $hkcu -Name 'NETDIR') { Test-Path -LiteralPath (Normalize-PathString (Get-RegistryValueSafe -Path $hkcu -Name 'NETDIR')) } else { $false })
        HklmWowPath = $hklmWow
        HklmWowNetDir = Normalize-PathString (Get-RegistryValueSafe -Path $hklmWow -Name 'NETDIR')
        HklmWowNetDirAccessible = $(if (Get-RegistryValueSafe -Path $hklmWow -Name 'NETDIR') { Test-Path -LiteralPath (Normalize-PathString (Get-RegistryValueSafe -Path $hklmWow -Name 'NETDIR')) } else { $false })
        HklmPath = $hklm
        HklmNetDir = Normalize-PathString (Get-RegistryValueSafe -Path $hklm -Name 'NETDIR')
        HklmNetDirAccessible = $(if (Get-RegistryValueSafe -Path $hklm -Name 'NETDIR') { Test-Path -LiteralPath (Normalize-PathString (Get-RegistryValueSafe -Path $hklm -Name 'NETDIR')) } else { $false })
        CfgEntries = @($cfgEntries)
        EcgProcessRunning = @((Get-Process -Name 'ECGV6' -ErrorAction SilentlyContinue)).Count -gt 0
    }
}

function Build-RemediationPlan {
    param(
        $SourceSnapshot,
        [bool]$ApplyFixes = $false,
        [bool]$AuthorizationProvided = $false,
        [string[]]$AppliedChanges = @()
    )
    $planned = New-Object System.Collections.ArrayList
    $blocked = New-Object System.Collections.ArrayList
    $desiredNetDir = Normalize-PathString $SourceSnapshot.OfficialNetDir
    $regTargets = @(
        [PSCustomObject]@{ Name = 'HKCU'; Path = $SourceSnapshot.HkcuPath; Value = $SourceSnapshot.HkcuNetDir },
        [PSCustomObject]@{ Name = 'HKLM_WOW6432Node'; Path = $SourceSnapshot.HklmWowPath; Value = $SourceSnapshot.HklmWowNetDir },
        [PSCustomObject]@{ Name = 'HKLM'; Path = $SourceSnapshot.HklmPath; Value = $SourceSnapshot.HklmNetDir }
    )
    foreach ($target in $regTargets) {
        if ((Normalize-PathString $target.Value) -ne $desiredNetDir) {
            [void]$planned.Add("Atualizar NETDIR em $($target.Name) para $desiredNetDir")
        }
    }
    foreach ($cfgEntry in @($SourceSnapshot.CfgEntries)) {
        if ((Normalize-PathString $cfgEntry.NetDir) -ne $desiredNetDir) {
            [void]$planned.Add("Atualizar IDAPI32.CFG em $($cfgEntry.Path) para $desiredNetDir")
        }
    }
    if (-not $SourceSnapshot.DesiredDbAccessible) { [void]$blocked.Add('Banco oficial configurado inacessivel no momento da decisao.') }
    if (-not $SourceSnapshot.DesiredNetDirAccessible) { [void]$blocked.Add('NetDir oficial configurado inacessivel no momento da decisao.') }
    if ($SourceSnapshot.EcgProcessRunning) { [void]$blocked.Add('ECGV6 aberto nesta estacao; remediacao segura exige aplicativo fechado.') }
    $needs = $planned.Count -gt 0
    $finalState = 'COMPLIANT'
    if ($needs -and $blocked.Count -gt 0) { $finalState = 'BLOCKED' }
    elseif ($needs -and $ApplyFixes -and -not $AuthorizationProvided) { $finalState = 'AUTH_REQUIRED' }
    elseif ($needs -and $ApplyFixes -and $AuthorizationProvided) { $finalState = 'READY_TO_APPLY' }
    elseif ($needs) { $finalState = 'DRIFT_DETECTED' }
    if ($AppliedChanges.Count -gt 0) { $finalState = 'REMEDIATED' }
    return [PSCustomObject]@{
        NeedsRemediation = $needs
        AuthorizationRequired = $needs
        AuthorizationProvided = $AuthorizationProvided
        ApplyRequested = $ApplyFixes
        PlannedActions = @($planned)
        Blockers = @($blocked)
        AppliedChanges = @($AppliedChanges)
        FinalState = $finalState
        Summary = $(if (-not $needs) { 'Sem drift acionavel; nenhuma mudanca planejada.' } elseif ($blocked.Count -gt 0) { 'Drift detectado, porem remediacao bloqueada pelos pre-checks.' } elseif ($ApplyFixes -and $AuthorizationProvided) { 'Drift detectado; remediacao autorizada e pronta para aplicar.' } elseif ($ApplyFixes) { 'Drift detectado; autorizacao explicita ausente para aplicar a remediacao.' } else { 'Drift detectado; relatorio em modo somente leitura com plano de correcao.' })
    }
}

function Get-BdeDriftSummary {
    param($SourceSnapshot)
    $desired = Normalize-PathString $SourceSnapshot.OfficialNetDir
    $drifts = New-Object System.Collections.ArrayList
    if ((Normalize-PathString $SourceSnapshot.HkcuNetDir) -and (Normalize-PathString $SourceSnapshot.HkcuNetDir) -ne $desired) { [void]$drifts.Add('HKCU') }
    if ((Normalize-PathString $SourceSnapshot.HklmWowNetDir) -and (Normalize-PathString $SourceSnapshot.HklmWowNetDir) -ne $desired) { [void]$drifts.Add('HKLM_WOW6432Node') }
    if ((Normalize-PathString $SourceSnapshot.HklmNetDir) -and (Normalize-PathString $SourceSnapshot.HklmNetDir) -ne $desired) { [void]$drifts.Add('HKLM') }
    foreach ($cfg in @($SourceSnapshot.CfgEntries)) {
        if ((Normalize-PathString $cfg.NetDir) -and (Normalize-PathString $cfg.NetDir) -ne $desired) { [void]$drifts.Add([System.IO.Path]::GetFileName($cfg.Path)) }
    }
    return @($drifts | Select-Object -Unique)
}


function Resolve-EcgPaths {
    param($Profile)
    $sourceSnapshot = Get-BdeSourceSnapshot -Profile $Profile
    $officialDb = Normalize-PathString $sourceSnapshot.OfficialDbPath
    $officialNetDir = Normalize-PathString $sourceSnapshot.OfficialNetDir
    $officialExe = Normalize-PathString $sourceSnapshot.OfficialExePath

    $effectiveDb = $officialDb
    $effectiveNetDir = $officialNetDir
    $dbSource = if (Test-IsUncPath $officialDb) { 'UNC oficial (INI)' } else { 'INI/valor oficial' }
    $netDirSource = if (Test-IsUncPath $officialNetDir) { 'UNC oficial (INI)' } else { 'INI/valor oficial' }

    if ($sourceSnapshot.HwDb -and $sourceSnapshot.HwDbAccessible -and -not (Test-Path -LiteralPath $effectiveDb)) {
        $effectiveDb = $sourceSnapshot.HwDb
        $dbSource = 'Variavel HW_CAMINHO_DB'
    }

    $preferredSources = @()
    if ($sourceSnapshot.HkcuNetDirAccessible) { $preferredSources += [PSCustomObject]@{ Value = $sourceSnapshot.HkcuNetDir; Source = 'Registro BDE (HKCU)' } }
    if ($sourceSnapshot.HklmWowNetDirAccessible) { $preferredSources += [PSCustomObject]@{ Value = $sourceSnapshot.HklmWowNetDir; Source = 'Registro BDE (HKLM WOW6432Node)' } }
    if ($sourceSnapshot.HklmNetDirAccessible) { $preferredSources += [PSCustomObject]@{ Value = $sourceSnapshot.HklmNetDir; Source = 'Registro BDE (HKLM)' } }
    foreach ($cfgEntry in @($sourceSnapshot.CfgEntries | Where-Object { $_.Accessible })) {
        $preferredSources += [PSCustomObject]@{ Value = $cfgEntry.NetDir; Source = ('IDAPI32.CFG: ' + $cfgEntry.Path) }
    }
    foreach ($candidate in @($preferredSources)) {
        if ((Normalize-PathString $candidate.Value) -eq $officialNetDir) {
            $effectiveNetDir = Normalize-PathString $candidate.Value
            $netDirSource = $candidate.Source
            break
        }
    }
    if (-not (Test-Path -LiteralPath $effectiveNetDir)) {
        foreach ($candidate in @($preferredSources)) {
            if ($candidate.Value) {
                $effectiveNetDir = Normalize-PathString $candidate.Value
                $netDirSource = $candidate.Source
                break
            }
        }
    }
    if (-not (Test-Path -LiteralPath $effectiveNetDir) -and $sourceSnapshot.HwNetDirAccessible) {
        $effectiveNetDir = $sourceSnapshot.HwNetDir
        $netDirSource = 'Variavel HW_CAMINHO_DB'
    }

    return [PSCustomObject]@{
        ExePath = $officialExe
        ExeAccessible = (Test-Path -LiteralPath $officialExe)
        DatabasePath = $effectiveDb
        DatabasePathSource = $dbSource
        DatabaseAccessible = (Test-Path -LiteralPath $effectiveDb)
        NetDirPath = $effectiveNetDir
        NetDirSource = $netDirSource
        NetDirAccessible = (Test-Path -LiteralPath $effectiveNetDir)
        DesiredDbPath = $officialDb
        DesiredNetDir = $officialNetDir
        DesiredExePath = $officialExe
        DatabaseHost = Get-UncHostFromPath -Path $effectiveDb
        NetDirHost = Get-UncHostFromPath -Path $effectiveNetDir
        BdeSources = $sourceSnapshot
    }
}

function Get-TimelineSample {
    param(
        $MachineType,
        $Paths,
        $SampleIndex,
        [int]$ProcessCaptureThresholdPercent = 80,
        [int]$TopProcessCount = 3,
        [bool]$EnableLatencyMetrics = $true,
        [bool]$EnableEcgProcessMetrics = $true,
        [bool]$EnableDiskMetrics = $false,
        [bool]$EnableNetworkMetrics = $false
    )

    $sampleTime = Get-Date

    $dbProbeMs = $null
    $netDirProbeMs = $null
    if ($EnableLatencyMetrics) {
        $dbProbe = Measure-PathProbeSafe -Path $Paths.DatabasePath
        $netProbe = Measure-PathProbeSafe -Path $Paths.NetDirPath
        $dbOk = [bool]$dbProbe.Accessible
        $netDirOk = [bool]$netProbe.Accessible
        $dbProbeMs = $dbProbe.ProbeMs
        $netDirProbeMs = $netProbe.ProbeMs
    }
    else {
        $dbOk = Test-Path $Paths.DatabasePath
        $netDirOk = Test-Path $Paths.NetDirPath
    }

    $exeOk = Test-Path $Paths.ExePath
    $cpu = Get-CpuPercentSafe
    $lockCount = Get-LockFileCountSafe -NetDirPath $Paths.NetDirPath
    $smb = Get-SmbMetricsSafe -MachineType $MachineType -DbHost $Paths.DatabaseHost -NetDirHost $Paths.NetDirHost

    $ecgProcessCpuPercent = $null
    if ($EnableEcgProcessMetrics) {
        $ecgProcessCpuPercent = Get-ProcessCounterCpuPercentSafe -ProcessPattern 'ECGV6*'
    }

    $physicalDiskQueueLength = $null
    if ($EnableDiskMetrics) {
        $physicalDiskQueueLength = Get-PhysicalDiskQueueLengthSafe
    }

    $networkBytesTotalPerSec = $null
    if ($EnableNetworkMetrics) {
        $networkBytesTotalPerSec = Get-NetworkBytesTotalPerSecSafe
    }

    $shouldCaptureTopProcesses = $false
    if ($null -ne $cpu -and [double]$cpu -ge $ProcessCaptureThresholdPercent) { $shouldCaptureTopProcesses = $true }
    if ($null -ne $ecgProcessCpuPercent -and [double]$ecgProcessCpuPercent -ge $ProcessCaptureThresholdPercent) { $shouldCaptureTopProcesses = $true }

    $topCpuProcesses = @()
    if ($shouldCaptureTopProcesses) {
        $topCpuProcesses = @(Get-TopCpuProcessesSafe -Top $TopProcessCount)
    }

    return [PSCustomObject]@{
        SampleIndex = $SampleIndex
        Timestamp = $sampleTime.ToString('HH:mm:ss')
        CpuPercent = $cpu
        EcgProcessCpuPercent = $ecgProcessCpuPercent
        DatabaseAccessible = $dbOk
        NetDirAccessible = $netDirOk
        ExeAccessible = $exeOk
        DatabaseProbeMs = $dbProbeMs
        NetDirProbeMs = $netDirProbeMs
        LockFileCount = $lockCount
        SmbConnectionCount = $smb.SmbConnectionCount
        SmbSessionCount = $smb.SmbSessionCount
        RelevantOpenFileCount = $smb.RelevantOpenFileCount
        RelevantNetDirOpenFileCount = $smb.RelevantNetDirOpenFileCount
        SmbQueryTimedOut = $smb.SmbQueryTimedOut
        PhysicalDiskQueueLength = $physicalDiskQueueLength
        NetworkBytesTotalPerSec = $networkBytesTotalPerSec
        TopCpuProcesses = @($topCpuProcesses)
    }
}

function Collect-Timeline {
    param(
        $MachineType,
        $Paths,
        [int]$Minutes,
        [int]$IntervalSeconds,
        [int]$ProcessCaptureThresholdPercent = 80,
        [int]$TopProcessCount = 3,
        [bool]$EnableLatencyMetrics = $true,
        [bool]$EnableEcgProcessMetrics = $true,
        [bool]$EnableDiskMetrics = $false,
        [bool]$EnableNetworkMetrics = $false
    )

    $samples = @()
    $started = Get-Date
    $targetSamples = [Math]::Max(1, [Math]::Floor(($Minutes * 60) / $IntervalSeconds))
    for ($i = 1; $i -le $targetSamples; $i++) {
        $sample = Get-TimelineSample -MachineType $MachineType -Paths $Paths -SampleIndex $i -ProcessCaptureThresholdPercent $ProcessCaptureThresholdPercent -TopProcessCount $TopProcessCount -EnableLatencyMetrics $EnableLatencyMetrics -EnableEcgProcessMetrics $EnableEcgProcessMetrics -EnableDiskMetrics $EnableDiskMetrics -EnableNetworkMetrics $EnableNetworkMetrics
        $samples += $sample

        $dominantProc = $null
        if ($sample.TopCpuProcesses -and @($sample.TopCpuProcesses).Count -gt 0) {
            $dominantProc = Get-CpuResponsibleProcessSummary -ProcessInfo @($sample.TopCpuProcesses)[0] -IncludeOwner:$false -IncludeCpu:$true
        }

        $logParts = New-Object System.Collections.Generic.List[string]
        [void]$logParts.Add(("Amostra {0}/{1}" -f $i, $targetSamples))
        [void]$logParts.Add(("CPU={0}%" -f $sample.CpuPercent))
        if ($null -ne $sample.EcgProcessCpuPercent) { [void]$logParts.Add(("ECGV6={0}%" -f $sample.EcgProcessCpuPercent)) }
        [void]$logParts.Add(("Locks={0}" -f $sample.LockFileCount))
        [void]$logParts.Add(("DB={0}" -f $sample.DatabaseAccessible))
        [void]$logParts.Add(("NetDir={0}" -f $sample.NetDirAccessible))
        if ($null -ne $sample.DatabaseProbeMs) { [void]$logParts.Add(("DBms={0}" -f $sample.DatabaseProbeMs)) }
        if ($null -ne $sample.NetDirProbeMs) { [void]$logParts.Add(("NetMs={0}" -f $sample.NetDirProbeMs)) }
        if ($null -ne $sample.PhysicalDiskQueueLength) { [void]$logParts.Add(("DiskQ={0}" -f $sample.PhysicalDiskQueueLength)) }
        if ($null -ne $sample.NetworkBytesTotalPerSec) { [void]$logParts.Add(("NetBytes={0}" -f $sample.NetworkBytesTotalPerSec)) }
        if (-not [string]::IsNullOrWhiteSpace($dominantProc)) { [void]$logParts.Add(("TopProc={0}" -f $dominantProc)) }

        Log ($logParts -join ' | ')
        if ($i -lt $targetSamples) { Start-Sleep -Seconds $IntervalSeconds }
    }
    $ended = Get-Date
    return [PSCustomObject]@{
        StartedAt = $started.ToString('dd/MM/yyyy HH:mm:ss')
        EndedAt = $ended.ToString('dd/MM/yyyy HH:mm:ss')
        ObservationMinutes = $Minutes
        IntervalSeconds = $IntervalSeconds
        SampleCount = $samples.Count
        Samples = $samples
    }
}

function Build-PassiveBenchmark {
    param($Timeline, $StagePriority, $SymptomCode)
    $samples = @($Timeline.Samples)
    $sampleCount = $samples.Count

    $cpuStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.CpuPercent })
    $ecgCpuStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.EcgProcessCpuPercent })
    $lockStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.LockFileCount })
    $dbProbeStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.DatabaseProbeMs })
    $netProbeStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.NetDirProbeMs })
    $diskQueueStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.PhysicalDiskQueueLength })
    $networkBytesStats = Get-ArrayStats -Values @($samples | ForEach-Object { $_.NetworkBytesTotalPerSec })

    $avgCpu = $cpuStats.Average
    $maxCpu = $cpuStats.Maximum
    $avgEcgCpu = $ecgCpuStats.Average
    $maxEcgCpu = $ecgCpuStats.Maximum
    $peakLocks = $lockStats.Maximum
    $minLocks = $lockStats.Minimum
    $lockAverage = $lockStats.Average

    $avgDbProbeMs = $dbProbeStats.Average
    $p95DbProbeMs = Get-PercentileSafe -Values @($samples | ForEach-Object { $_.DatabaseProbeMs }) -Percentile 95
    $maxDbProbeMs = $dbProbeStats.Maximum
    $avgNetProbeMs = $netProbeStats.Average
    $p95NetProbeMs = Get-PercentileSafe -Values @($samples | ForEach-Object { $_.NetDirProbeMs }) -Percentile 95
    $maxNetProbeMs = $netProbeStats.Maximum

    $avgDiskQueueLength = $diskQueueStats.Average
    $peakDiskQueueLength = $diskQueueStats.Maximum
    $avgNetworkBytesTotalPerSec = $networkBytesStats.Average
    $peakNetworkBytesTotalPerSec = $networkBytesStats.Maximum

    $dbUnavailable = @($samples | Where-Object { -not $_.DatabaseAccessible }).Count
    $netUnavailable = @($samples | Where-Object { -not $_.NetDirAccessible }).Count
    $smbTimeout = @($samples | Where-Object { $_.SmbQueryTimedOut }).Count
    $cpuBurstSamples90Plus = @($samples | Where-Object { $null -ne $_.CpuPercent -and [double]$_.CpuPercent -ge 90 }).Count
    $latencyWarningSamples = @($samples | Where-Object {
        (($null -ne $_.DatabaseProbeMs) -and [double]$_.DatabaseProbeMs -ge 100) -or
        (($null -ne $_.NetDirProbeMs) -and [double]$_.NetDirProbeMs -ge 100)
    }).Count
    $latencyCriticalSamples = @($samples | Where-Object {
        (($null -ne $_.DatabaseProbeMs) -and [double]$_.DatabaseProbeMs -ge 250) -or
        (($null -ne $_.NetDirProbeMs) -and [double]$_.NetDirProbeMs -ge 250)
    }).Count

    $samplesWithProcessCapture = @($samples | Where-Object { $null -ne $_.TopCpuProcesses -and @($_.TopCpuProcesses).Count -gt 0 }).Count
    $peakCpuProcessSample = $null
    $peakCpuResponsibleProcess = $null
    $peakCpuResponsibleSummary = $null
    if ($samplesWithProcessCapture -gt 0) {
        $peakCpuProcessSample = @($samples | Where-Object {
            $null -ne $_.TopCpuProcesses -and @($_.TopCpuProcesses).Count -gt 0
        } | Sort-Object -Property @{ Expression = { if ($null -ne $_.CpuPercent) { [double]$_.CpuPercent } else { -1 } }; Descending = $true }, @{ Expression = { $_.SampleIndex }; Descending = $false } | Select-Object -First 1)
        if ($peakCpuProcessSample.Count -gt 0) {
            $peakCpuProcessSample = $peakCpuProcessSample[0]
            $peakCpuResponsibleProcess = @($peakCpuProcessSample.TopCpuProcesses)[0]
            $peakCpuResponsibleSummary = Get-CpuResponsibleProcessSummary -ProcessInfo $peakCpuResponsibleProcess -IncludeOwner:$true -IncludeCpu:$true
        }
    }

    $lockValues = New-Object System.Collections.Generic.List[string]
    foreach ($sample in $samples) {
        if ($null -ne $sample.LockFileCount) {
            $lockKey = [string]([int]$sample.LockFileCount)
            if (-not $lockValues.Contains($lockKey)) { [void]$lockValues.Add($lockKey) }
        }
    }
    $lockUniqueValuesCount = $lockValues.Count
    $lockSpread = $null
    if ($null -ne $peakLocks -and $null -ne $minLocks) {
        $lockSpread = [math]::Abs([double]$peakLocks - [double]$minLocks)
    }

    $nominalLockBaseline = $false
    if ($null -ne $peakLocks -and $peakLocks -ge 1) {
        if ($peakLocks -le 2 -and $lockUniqueValuesCount -le 1 -and $dbUnavailable -eq 0 -and $netUnavailable -eq 0 -and $smbTimeout -eq 0) {
            $nominalLockBaseline = $true
        }
    }

    $severity = 0
    if ($dbUnavailable -gt 0) { $severity += [math]::Min(40, $dbUnavailable * 10) }
    if ($netUnavailable -gt 0) { $severity += [math]::Min(35, $netUnavailable * 10) }

    if ($null -ne $avgDbProbeMs) {
        if ($avgDbProbeMs -ge 250) { $severity += 8 }
        elseif ($avgDbProbeMs -ge 100) { $severity += 4 }
    }
    if ($null -ne $avgNetProbeMs) {
        if ($avgNetProbeMs -ge 250) { $severity += 8 }
        elseif ($avgNetProbeMs -ge 100) { $severity += 4 }
    }
    if ($null -ne $p95DbProbeMs) {
        if ($p95DbProbeMs -ge 250) { $severity += 6 }
        elseif ($p95DbProbeMs -ge 100) { $severity += 3 }
    }
    if ($null -ne $p95NetProbeMs) {
        if ($p95NetProbeMs -ge 250) { $severity += 6 }
        elseif ($p95NetProbeMs -ge 100) { $severity += 3 }
    }
    if ($latencyCriticalSamples -ge 2) { $severity += [math]::Min(10, $latencyCriticalSamples * 2) }
    elseif ($latencyWarningSamples -ge 3) { $severity += [math]::Min(6, $latencyWarningSamples) }

    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 85) { $severity += 25 }
        elseif ($avgCpu -ge 70) { $severity += 15 }
        elseif ($avgCpu -ge 55) { $severity += 5 }
    }

    if ($null -ne $maxCpu) {
        if ($maxCpu -ge 95) { $severity += 8 }
        elseif ($maxCpu -ge 80) { $severity += 4 }
    }

    if ($cpuBurstSamples90Plus -ge 3) {
        $severity += [math]::Min(12, ($cpuBurstSamples90Plus * 2))
    }

    if ($null -ne $avgEcgCpu) {
        if ($avgEcgCpu -ge 70) { $severity += 6 }
        elseif ($avgEcgCpu -ge 40) { $severity += 3 }
    }
    if ($null -ne $maxEcgCpu) {
        if ($maxEcgCpu -ge 85) { $severity += 6 }
        elseif ($maxEcgCpu -ge 60) { $severity += 3 }
    }

    if ($null -ne $peakDiskQueueLength) {
        if ($peakDiskQueueLength -ge 8) { $severity += 5 }
        elseif ($peakDiskQueueLength -ge 3) { $severity += 2 }
    }

    if ($null -ne $peakLocks) {
        if ($peakLocks -ge 6) { $severity += 18 }
        elseif ($peakLocks -ge 4) { $severity += 12 }
        elseif ($peakLocks -ge 3) { $severity += 6 }
        elseif ($peakLocks -ge 1 -and -not $nominalLockBaseline) { $severity += 3 }
    }

    if ($smbTimeout -gt 0) { $severity += [math]::Min(10, $smbTimeout * 3) }

    $severity = [math]::Min(100, $severity)
    $pressure = if ($severity -ge 60) { 'Pressionado' } elseif ($severity -ge 30) { 'Atencao' } else { 'Estavel' }

    return [PSCustomObject]@{
        Mode = 'Sem interacao'
        PressureLabel = $pressure
        SeverityScore = $severity
        SampleCount = $sampleCount
        AverageCpuPercent = $avgCpu
        PeakCpuPercent = $maxCpu
        AverageEcgProcessCpuPercent = $avgEcgCpu
        PeakEcgProcessCpuPercent = $maxEcgCpu
        PeakLockFileCount = $peakLocks
        MinimumLockFileCount = $minLocks
        AverageLockFileCount = $lockAverage
        LockUniqueValuesCount = $lockUniqueValuesCount
        LockSpread = $lockSpread
        NominalLockBaselineLikely = $nominalLockBaseline
        CpuBurstSamples90Plus = $cpuBurstSamples90Plus
        AverageDatabaseProbeMs = $avgDbProbeMs
        P95DatabaseProbeMs = $p95DbProbeMs
        PeakDatabaseProbeMs = $maxDbProbeMs
        AverageNetDirProbeMs = $avgNetProbeMs
        P95NetDirProbeMs = $p95NetProbeMs
        PeakNetDirProbeMs = $maxNetProbeMs
        LatencyWarningSamples = $latencyWarningSamples
        LatencyCriticalSamples = $latencyCriticalSamples
        AverageDiskQueueLength = $avgDiskQueueLength
        PeakDiskQueueLength = $peakDiskQueueLength
        AverageNetworkBytesTotalPerSec = $avgNetworkBytesTotalPerSec
        PeakNetworkBytesTotalPerSec = $peakNetworkBytesTotalPerSec
        SamplesWithProcessCapture = $samplesWithProcessCapture
        PeakCpuResponsibleProcess = if ($peakCpuResponsibleProcess) { [string]$peakCpuResponsibleProcess.Name } else { $null }
        PeakCpuResponsiblePid = if ($peakCpuResponsibleProcess) { $peakCpuResponsibleProcess.ProcessId } else { $null }
        PeakCpuResponsibleOwner = if ($peakCpuResponsibleProcess) { $peakCpuResponsibleProcess.Owner } else { $null }
        PeakCpuResponsibleProcessCpuPercent = if ($peakCpuResponsibleProcess) { $peakCpuResponsibleProcess.CpuPercent } else { $null }
        PeakCpuResponsibleIsEcgRelated = if ($peakCpuResponsibleProcess) { [bool]$peakCpuResponsibleProcess.IsEcgRelated } else { $null }
        PeakCpuResponsiblePath = if ($peakCpuResponsibleProcess) { $peakCpuResponsibleProcess.Path } else { $null }
        PeakCpuResponsibleSummary = $peakCpuResponsibleSummary
        PeakCpuResponsibleSampleTimestamp = if ($peakCpuProcessSample) { $peakCpuProcessSample.Timestamp } else { $null }
        PeakCpuTopProcesses = if ($peakCpuProcessSample) { @($peakCpuProcessSample.TopCpuProcesses) } else { @() }
        DatabaseUnavailableSamples = $dbUnavailable
        NetDirUnavailableSamples = $netUnavailable
        SmbTimeoutSamples = $smbTimeout
    }
}

function Build-AnalysisModel {
    param($MachineInfo, $Paths, $Timeline, $PassiveBenchmark, $StagePriority, $SymptomCode)

    $avgCpu = $PassiveBenchmark.AverageCpuPercent
    $peakCpu = $PassiveBenchmark.PeakCpuPercent
    $avgEcgCpu = $PassiveBenchmark.AverageEcgProcessCpuPercent
    $peakEcgCpu = $PassiveBenchmark.PeakEcgProcessCpuPercent
    $peakLocks = $PassiveBenchmark.PeakLockFileCount
    $dbUnavail = $PassiveBenchmark.DatabaseUnavailableSamples
    $netUnavail = $PassiveBenchmark.NetDirUnavailableSamples
    $smbTimeout = $PassiveBenchmark.SmbTimeoutSamples
    $severity = $PassiveBenchmark.SeverityScore
    $nominalLockBaseline = [bool]$PassiveBenchmark.NominalLockBaselineLikely
    $lockSpread = $PassiveBenchmark.LockSpread
    $cpuBurstSamples90Plus = $PassiveBenchmark.CpuBurstSamples90Plus
    $responsibleProcessSummary = [string]$PassiveBenchmark.PeakCpuResponsibleSummary
    $responsibleProcessName = [string]$PassiveBenchmark.PeakCpuResponsibleProcess
    $responsibleProcessIsEcg = [bool]$PassiveBenchmark.PeakCpuResponsibleIsEcgRelated
    $avgDbProbeMs = $PassiveBenchmark.AverageDatabaseProbeMs
    $p95DbProbeMs = $PassiveBenchmark.P95DatabaseProbeMs
    $avgNetProbeMs = $PassiveBenchmark.AverageNetDirProbeMs
    $p95NetProbeMs = $PassiveBenchmark.P95NetDirProbeMs
    $latencyWarningSamples = $PassiveBenchmark.LatencyWarningSamples
    $latencyCriticalSamples = $PassiveBenchmark.LatencyCriticalSamples
    $peakDiskQueueLength = $PassiveBenchmark.PeakDiskQueueLength

    $findings = @()
    $hypothesisSupport = @()
    $counterEvidence = @()
    $inconclusive = @()
    $discarded = @()
    $scores = @{ LOCAL = 0; SHARE = 0; LOCK = 0; SOFTWARE = 0 }

    if (-not $Paths.ExeAccessible) {
        $findings += 'Executavel oficial do ECG nao foi localizado no caminho padrao da ferramenta.'
        $scores.LOCAL += 4
        $hypothesisSupport += 'Executavel oficial inacessivel na estacao durante a rodada.'
    }
    else {
        $discarded += 'Executavel oficial localizado no caminho padrao.'
        $counterEvidence += 'Executavel oficial acessivel no caminho padrao.'
    }

    if (-not $Paths.DatabaseAccessible) {
        $findings += 'Banco do ECG inacessivel no caminho efetivo da rodada.'
        $scores.SHARE += 5
        $hypothesisSupport += 'Banco inacessivel no caminho efetivo durante a rodada.'
    }
    else {
        $discarded += 'Banco do ECG acessivel no caminho efetivo da rodada.'
        $counterEvidence += 'Banco acessivel no caminho efetivo nesta rodada.'
    }

    if (-not $Paths.NetDirAccessible) {
        $findings += 'NetDir inacessivel no caminho efetivo da rodada.'
        $scores.SHARE += 5
        $hypothesisSupport += 'NetDir inacessivel no caminho efetivo durante a rodada.'
    }
    else {
        $discarded += 'NetDir acessivel no caminho efetivo da rodada.'
        $counterEvidence += 'NetDir acessivel no caminho efetivo nesta rodada.'
    }

    $dbPathMatchesDesired = (Normalize-PathString $Paths.DatabasePath) -eq (Normalize-PathString $Paths.DesiredDbPath)
    if (-not $dbPathMatchesDesired) {
        $findings += 'A rodada nao conseguiu operar com o caminho oficial do banco; foi necessario caminho alternativo.'
        $scores.LOCAL += 3
        $hypothesisSupport += 'Houve fallback de caminho para o banco, divergente do caminho oficial configurado.'
    }
    else {
        $discarded += 'Banco operando a partir do caminho oficial configurado nesta rodada.'
        $counterEvidence += 'Banco operando no caminho oficial configurado, sem fallback.'
    }

    $netDirMatchesDesired = (Normalize-PathString $Paths.NetDirPath) -eq (Normalize-PathString $Paths.DesiredNetDir)
    $sourceDrifts = @()
    if ($null -ne $Paths.BdeSources) { $sourceDrifts = @(Get-BdeDriftSummary -SourceSnapshot $Paths.BdeSources) }
    if (-not $netDirMatchesDesired) {
        $findings += 'NetDir efetivo divergente do caminho oficial configurado.'
        $scores.LOCAL += 2
        $hypothesisSupport += 'NetDir efetivo divergente do caminho oficial configurado.'
    }
    else {
        $discarded += 'NetDir operando no caminho oficial configurado nesta rodada.'
        $counterEvidence += 'NetDir operando no caminho oficial configurado.'
    }

    if ($sourceDrifts.Count -gt 0) {
        $inconclusive += ('Ha drift entre as camadas de configuracao do BDE: ' + (($sourceDrifts -join ', ')))
        $hypothesisSupport += 'Camadas de configuracao do BDE divergentes entre arquivo CFG e/ou registro.'
    }

    if ($dbUnavail -gt 0) {
        $findings += "Houve indisponibilidade do banco em $dbUnavail amostra(s) da rodada."
        $scores.SHARE += [math]::Min(4, $dbUnavail)
        $hypothesisSupport += "Banco instavel em $dbUnavail amostra(s) da rodada."
    }
    else {
        $discarded += 'Sem indisponibilidade observada do banco durante a janela coletada.'
        $counterEvidence += 'Sem indisponibilidade de banco na janela observada.'
    }

    if ($netUnavail -gt 0) {
        $findings += "Houve indisponibilidade do NetDir em $netUnavail amostra(s) da rodada."
        $scores.SHARE += [math]::Min(4, $netUnavail)
        $hypothesisSupport += "NetDir instavel em $netUnavail amostra(s) da rodada."
    }
    else {
        $discarded += 'Sem indisponibilidade observada do NetDir durante a janela coletada.'
        $counterEvidence += 'Sem indisponibilidade de NetDir na janela observada.'
    }

    $shareLatencyStrong = $false
    if (($null -ne $p95DbProbeMs -and $p95DbProbeMs -ge 100) -or ($null -ne $p95NetProbeMs -and $p95NetProbeMs -ge 100) -or $latencyWarningSamples -gt 0) {
        if (($null -ne $p95DbProbeMs -and $p95DbProbeMs -ge 250) -or ($null -ne $p95NetProbeMs -and $p95NetProbeMs -ge 250) -or $latencyCriticalSamples -ge 2) {
            $shareLatencyStrong = $true
            $scores.SHARE += 4
            $findings += "Compartilhamento acessivel, porem com latencia elevada na rodada (DB p95: $p95DbProbeMs ms | NetDir p95: $p95NetProbeMs ms)."
            $hypothesisSupport += 'Caminho compartilhado acessivel, porem lento na janela observada.'
        }
        else {
            $scores.SHARE += 2
            $inconclusive += "Houve aumento de latencia no compartilhamento sem indisponibilidade (DB p95: $p95DbProbeMs ms | NetDir p95: $p95NetProbeMs ms)."
        }
    }
    elseif (($null -ne $avgDbProbeMs) -or ($null -ne $avgNetProbeMs)) {
        $counterEvidence += "Latencia de acesso ao compartilhamento em faixa saudavel (DB media: $avgDbProbeMs ms | NetDir media: $avgNetProbeMs ms)."
    }

    $lockContentionStrong = $false
    if ($null -ne $peakLocks -and $peakLocks -ge 1) {
        if ($nominalLockBaseline) {
            $discarded += "Baseline de lock/controle compativel com operacao compartilhada observado no NetDir (pico estavel: $peakLocks)."
            $counterEvidence += "Locks em baseline nominal para ambiente compartilhado (pico estavel: $peakLocks), sem indicio adicional de incidente."
        }
        elseif ($peakLocks -ge 4 -or ($peakLocks -ge 3 -and ($dbUnavail -gt 0 -or $netUnavail -gt 0 -or $smbTimeout -gt 0)) -or ($null -ne $lockSpread -and $lockSpread -ge 2)) {
            $lockContentionStrong = $true
            $findings += "Atividade de lock/controle acima do baseline foi observada no NetDir durante a janela (pico: $peakLocks)."
            $scores.LOCK += [math]::Min(6, [int]$peakLocks)
            $hypothesisSupport += "Padrao de lock acima do baseline do NetDir (pico: $peakLocks)."
        }
        else {
            $inconclusive += "Locks/arquivos de controle foram vistos no NetDir (pico: $peakLocks), mas o sinal isolado ainda nao sustenta contencao por si so."
        }
    }
    else {
        $discarded += 'Sem lock/controle relevante observado no NetDir durante a janela.'
        $counterEvidence += 'Sem lock/controle relevante no NetDir durante a janela observada.'
    }

    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 70) {
            $findings += "CPU media elevada na rodada ($avgCpu%)."
            $scores.SOFTWARE += 4
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $hypothesisSupport += "CPU media elevada com processo dominante no pico: $responsibleProcessSummary."
            }
            else {
                $hypothesisSupport += "CPU media elevada na rodada ($avgCpu%)."
            }
        }
        elseif ($avgCpu -lt 55) {
            $discarded += "CPU media da rodada sem pressao sustentada relevante ($avgCpu%)."
            $counterEvidence += "CPU media sem pressao sustentada relevante ($avgCpu%)."
        }
        else {
            $inconclusive += 'CPU media ficou em faixa intermediaria; o sinal isolado nao fecha hipotese por si so.'
        }
    }

    if ($null -ne $avgEcgCpu) {
        if ($avgEcgCpu -ge 40) {
            $findings += "Processo ECGV6 com uso relevante de CPU na rodada (media: $avgEcgCpu% | pico: $peakEcgCpu%)."
            $scores.SOFTWARE += 3
            $hypothesisSupport += 'Uso elevado de CPU associado diretamente ao processo ECGV6.'
        }
        elseif ($avgEcgCpu -lt 15) {
            $counterEvidence += "Processo ECGV6 sem pressao sustentada relevante de CPU (media: $avgEcgCpu%)."
        }
    }

    if ($null -ne $peakCpu) {
        if ($peakCpu -ge 95 -and $cpuBurstSamples90Plus -ge 3) {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $findings += "Burst local de CPU relevante observado na estacao (pico: $peakCpu%). Processo dominante no pico: $responsibleProcessSummary."
                $hypothesisSupport += "Burst de CPU local relevante com processo dominante identificado: $responsibleProcessSummary."
            }
            else {
                $findings += "Burst local de CPU relevante observado na estacao (pico: $peakCpu%)."
                $hypothesisSupport += "Burst de CPU local relevante observado na estacao (pico: $peakCpu%)."
            }
            $scores.SOFTWARE += 4
        }
        elseif ($peakCpu -ge 80 -and $cpuBurstSamples90Plus -lt 3) {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $inconclusive += "Houve pico isolado de CPU ($peakCpu%), porem sem persistencia suficiente para fechar hipotese local. Processo dominante no pico: $responsibleProcessSummary."
            }
            else {
                $inconclusive += "Houve pico isolado de CPU ($peakCpu%), porem sem persistencia suficiente para fechar hipotese local."
            }
        }
    }

    if ($null -ne $peakDiskQueueLength) {
        if ($peakDiskQueueLength -ge 8) {
            $findings += "Fila de disco elevada observada na rodada (pico: $peakDiskQueueLength)."
            $scores.SOFTWARE += 2
            $hypothesisSupport += 'Sinal de pressao de disco local na estacao.'
        }
        elseif ($peakDiskQueueLength -ge 3) {
            $inconclusive += "Fila de disco moderada observada na rodada (pico: $peakDiskQueueLength)."
        }
    }

    if ($smbTimeout -gt 0) {
        $findings += 'Parte das consultas SMB excedeu timeout; leitura de compartilhamento deve ser interpretada com cautela.'
        $scores.SHARE += [math]::Min(3, $smbTimeout)
        $inconclusive += 'Consultas SMB com timeout reduzem a forca de alguns descartes de compartilhamento.'
    }
    else {
        $discarded += 'Sem timeout observado nas consultas SMB executadas pela ferramenta.'
        $counterEvidence += 'Sem timeout observado nas consultas SMB executadas pela ferramenta.'
    }

    $primary = 'SHARE'
    $maxScore = -1
    foreach ($k in $scores.Keys) {
        if ($scores[$k] -gt $maxScore) { $maxScore = $scores[$k]; $primary = $k }
    }
    if ($maxScore -le 0) { $primary = 'OK' }

    $hasStrongLocalBurst = ($null -ne $peakCpu -and $peakCpu -ge 95 -and $cpuBurstSamples90Plus -ge 3)
    $status = if ($maxScore -ge 8 -or (-not $Paths.ExeAccessible) -or ($dbUnavail -ge 2) -or ($netUnavail -ge 2) -or $latencyCriticalSamples -ge 3) {
        'CRITICO'
    }
    elseif ($maxScore -ge 4 -or ($null -ne $avgCpu -and $avgCpu -ge 65) -or $hasStrongLocalBurst -or $lockContentionStrong -or $shareLatencyStrong) {
        'LENTO'
    }
    elseif ($maxScore -eq 0 -and $Paths.ExeAccessible -and $Paths.DatabaseAccessible -and $Paths.NetDirAccessible -and $smbTimeout -eq 0) {
        'NORMAL'
    }
    else {
        'INCONCLUSIVO'
    }

    $confidence = 'Baixa'
    if ($status -eq 'NORMAL' -and $maxScore -eq 0) { $confidence = 'Media' }
    elseif ($maxScore -ge 8 -and $hypothesisSupport.Count -ge 2) { $confidence = 'Alta' }
    elseif ($maxScore -ge 4 -and $hypothesisSupport.Count -ge 1) { $confidence = 'Media' }

    $hypothesisMap = @{ LOCAL = 'Configuracao local'; SHARE = 'Compartilhamento/acesso'; LOCK = 'Contencao/lock'; SOFTWARE = 'Software/arquivo'; OK = 'Ambiente integro' }
    $primaryHypothesis = $hypothesisMap[$primary]
    $impactScope = switch ($primary) {
        'SHARE' { 'Sistema compartilhado' }
        'LOCK'  { 'Sistema compartilhado' }
        'OK'    { 'Sem impacto relevante observado' }
        default { 'Somente este computador' }
    }

    $recommendedAction = switch ($primary) {
        'SHARE' { 'Validar latencia do compartilhamento, tempo de resposta do UNC e permissoes antes de atuar no software.' }
        'LOCAL' { 'Conferir configuracao local do ECG/BDE, mapeamentos e aderencia aos caminhos oficiais da aplicacao.' }
        'LOCK' { 'Repetir a rodada durante o sintoma e revisar concorrencia de acesso, locks e fluxo de gravacao no NetDir.' }
        'SOFTWARE' {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessName)) {
                if ($responsibleProcessIsEcg) {
                    "Revisar comportamento local do ECG, priorizando o processo $responsibleProcessName e correlacionando com o momento do pico de CPU."
                }
                else {
                    "Revisar a estacao e o processo dominante no pico de CPU ($responsibleProcessName), verificando antivirus/EDR, tarefas concorrentes e interferencia externa ao ECG."
                }
            }
            else {
                'Revisar saude do software do ECG e comportamento local da estacao, priorizando processo do ECG, antivirus/EDR e competidores de CPU.'
            }
        }
        default {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                "Sem acao corretiva imediata. Repetir a coleta apenas durante o sintoma. Na rodada atual, o processo dominante no pico foi: $responsibleProcessSummary."
            }
            else {
                'Sem acao corretiva imediata. Repetir a coleta apenas durante o sintoma percebido.'
            }
        }
    }

    $summaryPhrase = "A rodada indica condicao $($status.ToLowerInvariant()) com hipotese principal em $($primaryHypothesis.ToLowerInvariant())."
    $probablePerception = switch ($status) {
        'CRITICO' { 'Usuario tende a perceber falha evidente, demora anormal ou impossibilidade pratica de continuar o fluxo do ECG.' }
        'LENTO' { 'Usuario tende a perceber lentidao real ou resposta inconsistente durante a janela observada.' }
        'NORMAL' { 'A rodada atual nao reuniu evidencia de degradacao ativa no backend compartilhado.' }
        default { 'Usuario pode ter percebido oscilacao anterior, mas a rodada atual nao reuniu evidencia forte de degradacao ativa.' }
    }

    $secondary = @()
    foreach ($k in $scores.Keys) {
        if ($k -ne $primary -and $scores[$k] -gt 0) {
            $secondary += [PSCustomObject]@{ Name = $hypothesisMap[$k]; Score = $scores[$k]; MainReasons = @() }
        }
    }

    return [PSCustomObject]@{
        Status = $status
        StatusCode = $status
        Confidence = $confidence
        PrimaryHypothesis = $primaryHypothesis
        ImpactScope = $impactScope
        SummaryPhrase = $summaryPhrase
        ProbablePerception = $probablePerception
        RecommendedAction = $recommendedAction
        Findings = $findings
        HypothesisSupport = $hypothesisSupport
        HypothesisCounterpoints = $counterEvidence
        InconclusivePoints = $inconclusive
        WhatDidNotIndicateFailure = $discarded
        SecondaryHypotheses = $secondary
        Metrics = [PSCustomObject]@{
            SeverityScore = $severity
            PressureLabel = $PassiveBenchmark.PressureLabel
            SampleCount = $Timeline.SampleCount
            AverageCpuPercent = $avgCpu
            PeakCpuPercent = $PassiveBenchmark.PeakCpuPercent
            AverageEcgProcessCpuPercent = $PassiveBenchmark.AverageEcgProcessCpuPercent
            PeakEcgProcessCpuPercent = $PassiveBenchmark.PeakEcgProcessCpuPercent
            PeakLockFileCount = $peakLocks
            MinimumLockFileCount = $PassiveBenchmark.MinimumLockFileCount
            AverageLockFileCount = $PassiveBenchmark.AverageLockFileCount
            LockUniqueValuesCount = $PassiveBenchmark.LockUniqueValuesCount
            LockSpread = $PassiveBenchmark.LockSpread
            NominalLockBaselineLikely = $PassiveBenchmark.NominalLockBaselineLikely
            CpuBurstSamples90Plus = $PassiveBenchmark.CpuBurstSamples90Plus
            AverageDatabaseProbeMs = $PassiveBenchmark.AverageDatabaseProbeMs
            P95DatabaseProbeMs = $PassiveBenchmark.P95DatabaseProbeMs
            PeakDatabaseProbeMs = $PassiveBenchmark.PeakDatabaseProbeMs
            AverageNetDirProbeMs = $PassiveBenchmark.AverageNetDirProbeMs
            P95NetDirProbeMs = $PassiveBenchmark.P95NetDirProbeMs
            PeakNetDirProbeMs = $PassiveBenchmark.PeakNetDirProbeMs
            LatencyWarningSamples = $PassiveBenchmark.LatencyWarningSamples
            LatencyCriticalSamples = $PassiveBenchmark.LatencyCriticalSamples
            AverageDiskQueueLength = $PassiveBenchmark.AverageDiskQueueLength
            PeakDiskQueueLength = $PassiveBenchmark.PeakDiskQueueLength
            AverageNetworkBytesTotalPerSec = $PassiveBenchmark.AverageNetworkBytesTotalPerSec
            PeakNetworkBytesTotalPerSec = $PassiveBenchmark.PeakNetworkBytesTotalPerSec
            SamplesWithProcessCapture = $PassiveBenchmark.SamplesWithProcessCapture
            PeakCpuResponsibleProcess = $PassiveBenchmark.PeakCpuResponsibleProcess
            PeakCpuResponsiblePid = $PassiveBenchmark.PeakCpuResponsiblePid
            PeakCpuResponsibleOwner = $PassiveBenchmark.PeakCpuResponsibleOwner
            PeakCpuResponsibleProcessCpuPercent = $PassiveBenchmark.PeakCpuResponsibleProcessCpuPercent
            PeakCpuResponsibleIsEcgRelated = $PassiveBenchmark.PeakCpuResponsibleIsEcgRelated
            PeakCpuResponsiblePath = $PassiveBenchmark.PeakCpuResponsiblePath
            PeakCpuResponsibleSummary = $PassiveBenchmark.PeakCpuResponsibleSummary
            PeakCpuResponsibleSampleTimestamp = $PassiveBenchmark.PeakCpuResponsibleSampleTimestamp
            DatabaseUnavailableSamples = $dbUnavail
            NetDirUnavailableSamples = $netUnavail
            SmbTimeoutSamples = $smbTimeout
        }
    }
}

function Get-ChartDefinition {
    param($Timeline)
    $samples = @($Timeline.Samples)
    $labels = @($samples | ForEach-Object { $_.Timestamp })
    $cpuValues = @($samples | ForEach-Object { if ($null -ne $_.CpuPercent) { [double]$_.CpuPercent } else { $null } })
    $lockValues = @($samples | ForEach-Object { if ($null -ne $_.LockFileCount) { [double]$_.LockFileCount } else { 0 } })
    $dbLatencyValues = @($samples | ForEach-Object { if ($null -ne $_.DatabaseProbeMs) { [double]$_.DatabaseProbeMs } else { $null } })
    $netLatencyValues = @($samples | ForEach-Object { if ($null -ne $_.NetDirProbeMs) { [double]$_.NetDirProbeMs } else { $null } })
    $ecgCpuValues = @($samples | ForEach-Object { if ($null -ne $_.EcgProcessCpuPercent) { [double]$_.EcgProcessCpuPercent } else { $null } })

    $lockStats = Get-ArrayStats -Values $lockValues
    $dbLatencyStats = Get-ArrayStats -Values $dbLatencyValues
    $netLatencyStats = Get-ArrayStats -Values $netLatencyValues
    $ecgCpuStats = Get-ArrayStats -Values $ecgCpuValues

    $maxLock = if ($null -ne $lockStats.Maximum -and $lockStats.Maximum -gt 0) { $lockStats.Maximum } else { 1 }
    $maxDbLatency = if ($null -ne $dbLatencyStats.Maximum -and $dbLatencyStats.Maximum -gt 0) { $dbLatencyStats.Maximum } else { 1 }
    $maxNetLatency = if ($null -ne $netLatencyStats.Maximum -and $netLatencyStats.Maximum -gt 0) { $netLatencyStats.Maximum } else { 1 }
    $maxEcgCpu = if ($null -ne $ecgCpuStats.Maximum -and $ecgCpuStats.Maximum -gt 0) { [math]::Max(100, $ecgCpuStats.Maximum) } else { 100 }

    return [PSCustomObject]@{
        labels = $labels
        series = @(
            [PSCustomObject]@{ name = 'CPU %'; values = $cpuValues; max = 100; color = '#2563eb' },
            [PSCustomObject]@{ name = 'Locks'; values = $lockValues; max = $maxLock; color = '#dc2626' },
            [PSCustomObject]@{ name = 'DB ms'; values = $dbLatencyValues; max = $maxDbLatency; color = '#7c3aed' },
            [PSCustomObject]@{ name = 'NetDir ms'; values = $netLatencyValues; max = $maxNetLatency; color = '#0f766e' },
            [PSCustomObject]@{ name = 'ECGV6 %'; values = $ecgCpuValues; max = $maxEcgCpu; color = '#ea580c' }
        )
    }
}

function Select-TopTextItems {
    param(
        [object[]]$Items,
        [int]$Max = 3,
        [string]$FallbackText = 'Sem evidencia adicional.'
    )
    $list = @()
    foreach ($item in @($Items)) {
        $text = [string]$item
        if (-not [string]::IsNullOrWhiteSpace($text)) { $list += $text.Trim() }
    }
    if ($list.Count -eq 0) { return @($FallbackText) }
    if ($list.Count -le $Max) { return @($list) }
    return @($list[0..($Max - 1)])
}

function Convert-TextItemsToHtmlList {
    param([object[]]$Items)
    return ((@($Items) | ForEach-Object { "<li>$([System.Net.WebUtility]::HtmlEncode([string]$_))</li>" }) -join '')
}

function Get-ProcessCaptureRows {
    param($Timeline)
    $rows = @()
    foreach ($sample in @($Timeline.Samples)) {
        $topItems = @($sample.TopCpuProcesses)
        if ($topItems.Count -eq 0) { continue }
        foreach ($item in $topItems) {
            $rows += [PSCustomObject]@{
                Timestamp = $sample.Timestamp
                SampleIndex = $sample.SampleIndex
                CpuPercent = $sample.CpuPercent
                ProcessName = $item.Name
                ProcessId = $item.ProcessId
                ProcessCpuPercent = $item.CpuPercent
                Owner = $item.Owner
                Path = $item.Path
                IsEcgRelated = $item.IsEcgRelated
            }
        }
    }
    return @($rows)
}

function Get-OperationalDecisionModel {
    param($Analysis, $Paths, $PassiveBenchmark)

    $backendStatus = 'OK'
    $backendLine = 'Banco, NetDir e SMB sem falha nesta rodada.'
    if (-not $Paths.DatabaseAccessible -or -not $Paths.NetDirAccessible) {
        $backendStatus = 'CRITICO'
        $backendLine = 'Ha indisponibilidade do caminho efetivo do banco ou do NetDir na estacao.'
    }
    elseif ($PassiveBenchmark.DatabaseUnavailableSamples -gt 0 -or $PassiveBenchmark.NetDirUnavailableSamples -gt 0 -or $PassiveBenchmark.SmbTimeoutSamples -gt 0) {
        $backendStatus = 'ATENCAO'
        $backendLine = 'Houve oscilacao de backend/SMB durante a janela; validar compartilhamento antes de atuar no software.'
    }
    elseif (($null -ne $PassiveBenchmark.P95DatabaseProbeMs -and $PassiveBenchmark.P95DatabaseProbeMs -ge 100) -or ($null -ne $PassiveBenchmark.P95NetDirProbeMs -and $PassiveBenchmark.P95NetDirProbeMs -ge 100)) {
        $backendStatus = 'ATENCAO'
        $backendLine = "Compartilhamento acessivel, porem com latencia elevada (DB p95: $($PassiveBenchmark.P95DatabaseProbeMs) ms | NetDir p95: $($PassiveBenchmark.P95NetDirProbeMs) ms)."
    }

    $responsibleProcessSummary = [string]$PassiveBenchmark.PeakCpuResponsibleSummary

    $localStatus = 'OK'
    $localLine = 'Sem pressao local sustentada relevante nesta rodada.'
    if (-not $Paths.ExeAccessible) {
        $localStatus = 'ATENCAO'
        $localLine = 'Executavel oficial nao foi localizado no caminho esperado nesta estacao.'
    }
    elseif (($null -ne $PassiveBenchmark.PeakCpuPercent -and $PassiveBenchmark.PeakCpuPercent -ge 95 -and $PassiveBenchmark.CpuBurstSamples90Plus -ge 3) -or ($null -ne $PassiveBenchmark.AverageCpuPercent -and $PassiveBenchmark.AverageCpuPercent -ge 70) -or ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent -and $PassiveBenchmark.PeakEcgProcessCpuPercent -ge 80)) {
        $localStatus = 'ATENCAO'
        if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
            $localLine = "Pressao local observada. Processo dominante no pico: $responsibleProcessSummary."
        }
        elseif ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent -and $PassiveBenchmark.PeakEcgProcessCpuPercent -ge 80) {
            $localLine = "ECGV6 com uso elevado de CPU na rodada (pico: $($PassiveBenchmark.PeakEcgProcessCpuPercent)%)."
        }
        elseif ($PassiveBenchmark.CpuBurstSamples90Plus -ge 3 -and $null -ne $PassiveBenchmark.PeakCpuPercent) {
            $localLine = "Burst local de CPU observado (pico: $($PassiveBenchmark.PeakCpuPercent)% em $($PassiveBenchmark.CpuBurstSamples90Plus) amostra(s))."
        }
        elseif ($null -ne $PassiveBenchmark.AverageCpuPercent) {
            $localLine = "CPU media local em faixa elevada nesta rodada ($($PassiveBenchmark.AverageCpuPercent)%)."
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
        $localLine = "Sem pressao local sustentada relevante nesta rodada. Processo dominante no pico isolado: $responsibleProcessSummary."
    }

    $lockLine = if ($PassiveBenchmark.NominalLockBaselineLikely) {
        "Locks em baseline nominal para ambiente compartilhado (pico estavel: $($PassiveBenchmark.PeakLockFileCount))."
    }
    elseif ($null -ne $PassiveBenchmark.PeakLockFileCount) {
        "Pico de lock observado: $($PassiveBenchmark.PeakLockFileCount)."
    }
    else {
        'Sem leitura relevante de lock nesta rodada.'
    }

    return [PSCustomObject]@{
        BackendStatus = $backendStatus
        BackendLine = $backendLine
        LocalStatus = $localStatus
        LocalLine = $localLine
        LockLine = $lockLine
    }
}

function Get-RelevantTimelineSamples {
    param($Timeline)
    $samples = @($Timeline.Samples)
    if ($samples.Count -eq 0) { return @() }

    $selected = New-Object System.Collections.ArrayList
    $previousLock = $null
    $previousConn = $null
    $previousDb = $null
    $previousNet = $null

    foreach ($sample in $samples) {
        $include = $false
        if ($sample.SampleIndex -eq 1 -or $sample.SampleIndex -eq $samples.Count) { $include = $true }
        if ($null -ne $sample.CpuPercent -and [double]$sample.CpuPercent -ge 80) { $include = $true }
        if ($null -ne $sample.EcgProcessCpuPercent -and [double]$sample.EcgProcessCpuPercent -ge 60) { $include = $true }
        if ($sample.SmbQueryTimedOut -or -not $sample.DatabaseAccessible -or -not $sample.NetDirAccessible) { $include = $true }
        if (($null -ne $sample.DatabaseProbeMs -and [double]$sample.DatabaseProbeMs -ge 100) -or ($null -ne $sample.NetDirProbeMs -and [double]$sample.NetDirProbeMs -ge 100)) { $include = $true }
        if ($null -ne $sample.PhysicalDiskQueueLength -and [double]$sample.PhysicalDiskQueueLength -ge 3) { $include = $true }
        if ($null -ne $previousLock -and $sample.LockFileCount -ne $previousLock) { $include = $true }
        if ($null -ne $previousConn -and $sample.SmbConnectionCount -ne $previousConn) { $include = $true }
        if ($null -ne $previousDb -and $sample.DatabaseAccessible -ne $previousDb) { $include = $true }
        if ($null -ne $previousNet -and $sample.NetDirAccessible -ne $previousNet) { $include = $true }
        if ($sample.TopCpuProcesses -and @($sample.TopCpuProcesses).Count -gt 0) { $include = $true }

        if ($include) {
            $already = @($selected | Where-Object { $_.SampleIndex -eq $sample.SampleIndex }).Count -gt 0
            if (-not $already) { [void]$selected.Add($sample) }
        }

        $previousLock = $sample.LockFileCount
        $previousConn = $sample.SmbConnectionCount
        $previousDb = $sample.DatabaseAccessible
        $previousNet = $sample.NetDirAccessible

        if ($selected.Count -ge 20) { break }
    }

    return @($selected)
}

function Build-HtmlReport {
    param($Analysis, $MachineInfo, $Paths, $Timeline, $PassiveBenchmark, $StagePriority, $SymptomCode, $Remediation = $null)
    $statusClass = switch ($Analysis.StatusCode) {
        'NORMAL' { 'normal' }
        'LENTO' { 'lento' }
        'CRITICO' { 'critico' }
        default { 'inconclusivo' }
    }

    $decision = Get-OperationalDecisionModel -Analysis $Analysis -Paths $Paths -PassiveBenchmark $PassiveBenchmark

    $topSupport = Select-TopTextItems -Items $Analysis.HypothesisSupport -Max 3 -FallbackText $(if ($Analysis.StatusCode -eq 'NORMAL') { 'Nenhum indicio relevante de falha ativa nesta rodada.' } else { 'Sem evidencia dominante suficiente nesta rodada.' })
    $topDiscard = Select-TopTextItems -Items $Analysis.WhatDidNotIndicateFailure -Max 5 -FallbackText 'Sem descarte adicional.'
    $topInconclusive = Select-TopTextItems -Items $Analysis.InconclusivePoints -Max 3 -FallbackText 'Sem ponto inconclusivo adicional.'
    $supportItems = Convert-TextItemsToHtmlList -Items $topSupport
    $discardItems = Convert-TextItemsToHtmlList -Items $topDiscard
    $inconclusiveItems = Convert-TextItemsToHtmlList -Items $topInconclusive

    $secondaryRows = if ($Analysis.SecondaryHypotheses.Count) {
        ($Analysis.SecondaryHypotheses | ForEach-Object {
            "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.Name))</td><td>$($_.Score)</td><td>Hipotese secundaria com evidencia parcial.</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="3">Sem hipotese secundaria relevante.</td></tr>'
    }

    $relevantSamples = @(Get-RelevantTimelineSamples -Timeline $Timeline)
    $relevantTimelineRows = if ($relevantSamples.Count -gt 0) {
        ($relevantSamples | ForEach-Object {
            $dominantProcess = if ($_.TopCpuProcesses -and @($_.TopCpuProcesses).Count -gt 0) { Get-CpuResponsibleProcessSummary -ProcessInfo @($_.TopCpuProcesses)[0] -IncludeOwner:$false -IncludeCpu:$true } else { 'N/D' }
            "<tr><td>$($_.Timestamp)</td><td>$($_.CpuPercent)</td><td>$(if ($null -ne $_.EcgProcessCpuPercent) { $_.EcgProcessCpuPercent } else { 'N/D' })</td><td>$($_.LockFileCount)</td><td>$(if ($null -ne $_.DatabaseProbeMs) { $_.DatabaseProbeMs } else { 'N/D' })</td><td>$(if ($null -ne $_.NetDirProbeMs) { $_.NetDirProbeMs } else { 'N/D' })</td><td>$(if ($null -ne $_.PhysicalDiskQueueLength) { $_.PhysicalDiskQueueLength } else { 'N/D' })</td><td>$(if ($null -ne $_.NetworkBytesTotalPerSec) { $_.NetworkBytesTotalPerSec } else { 'N/D' })</td><td>$(if ($null -ne $_.SmbConnectionCount) { $_.SmbConnectionCount } else { 'N/A' })</td><td>$(if ($_.DatabaseAccessible) { 'Sim' } else { 'Nao' })</td><td>$(if ($_.NetDirAccessible) { 'Sim' } else { 'Nao' })</td><td>$(if ($_.SmbQueryTimedOut) { 'Sim' } else { 'Nao' })</td><td>$([System.Net.WebUtility]::HtmlEncode($dominantProcess))</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="13">Sem eventos relevantes selecionados nesta rodada.</td></tr>'
    }

    $processCaptureRows = @(Get-ProcessCaptureRows -Timeline $Timeline)
    $processCaptureTableRows = if ($processCaptureRows.Count -gt 0) {
        ($processCaptureRows | ForEach-Object {
            "<tr><td>$($_.Timestamp)</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.ProcessName))</td><td>$($_.ProcessId)</td><td>$($_.CpuPercent)</td><td>$($_.ProcessCpuPercent)</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Owner))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($_.IsEcgRelated) { 'Sim' } else { 'Nao' })))</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="7">Nenhuma captura de processo dominante foi necessaria nesta rodada.</td></tr>'
    }

    $chartDef = Get-ChartDefinition -Timeline $Timeline
    $chartJson = $chartDef | ConvertTo-Json -Depth 10 -Compress
    $responsibleCpuText = if (-not [string]::IsNullOrWhiteSpace([string]$PassiveBenchmark.PeakCpuResponsibleSummary)) { [string]$PassiveBenchmark.PeakCpuResponsibleSummary } else { 'N/D' }

    $bdeSourceRows = ''
    if ($null -ne $Paths.BdeSources) {
        $cfgRows = @($Paths.BdeSources.CfgEntries | ForEach-Object {
            "<tr><td>IDAPI32.CFG</td><td>$([System.Net.WebUtility]::HtmlEncode($_.Path))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.NetDir))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($_.Accessible) { 'Sim' } else { 'Nao' })))</td></tr>"
        })
        $extraRows = @(
            "<tr><td>DesiredNetDir</td><td>INI oficial</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.DesiredNetDir))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($Paths.BdeSources.DesiredNetDirAccessible) { 'Sim' } else { 'Nao' })))</td></tr>",
            "<tr><td>HKCU</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HkcuPath))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HkcuNetDir))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($Paths.BdeSources.HkcuNetDirAccessible) { 'Sim' } else { 'Nao' })))</td></tr>",
            "<tr><td>HKLM WOW6432Node</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HklmWowPath))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HklmWowNetDir))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($Paths.BdeSources.HklmWowNetDirAccessible) { 'Sim' } else { 'Nao' })))</td></tr>",
            "<tr><td>HKLM</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HklmPath))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$Paths.BdeSources.HklmNetDir))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($Paths.BdeSources.HklmNetDirAccessible) { 'Sim' } else { 'Nao' })))</td></tr>"
        )
        $bdeSourceRows = (@($extraRows) + @($cfgRows)) -join ''
    }
    $remediationSummary = 'Nao informado.'
    $plannedActionsItems = '<li>Sem acao planejada.</li>'
    $blockerItems = '<li>Sem bloqueio ativo.</li>'
    if ($null -ne $Remediation) {
        $remediationSummary = [System.Net.WebUtility]::HtmlEncode([string]$Remediation.Summary)
        $plannedActionsItems = Convert-TextItemsToHtmlList -Items @($Remediation.PlannedActions) -FallbackText 'Sem acao planejada.'
        $blockerItems = Convert-TextItemsToHtmlList -Items @($Remediation.Blockers) -FallbackText 'Sem bloqueio ativo.'
    }

    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head><meta charset="UTF-8"><title>ECG Diagnostics Report</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f6fb;color:#1f2937}
.wrapper{max-width:1260px;margin:0 auto;padding:24px}
.card,details{background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(15,23,42,0.08);margin-bottom:16px}
.card{padding:20px}
details{padding:16px 20px}
summary{cursor:pointer;font-weight:700}
h1,h2,h3{margin-top:0}.badge{display:inline-block;padding:8px 14px;border-radius:999px;font-weight:700;font-size:13px}
.badge.normal{background:#dcfce7;color:#166534}.badge.lento{background:#fef3c7;color:#92400e}.badge.critico{background:#fee2e2;color:#991b1b}.badge.inconclusivo{background:#e5e7eb;color:#374151}
.grid{display:grid;gap:12px;grid-template-columns:repeat(2,1fr)}.grid-3{display:grid;gap:12px;grid-template-columns:repeat(3,1fr)}
.kv{border:1px solid #e5e7eb;border-radius:12px;padding:10px 12px;background:#fafbfd}.kv strong{display:block;font-size:12px;color:#6b7280}
.muted{color:#6b7280;font-size:12px}table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #e5e7eb;padding:8px 10px;text-align:left;vertical-align:top}th{background:#f9fafb}
canvas{width:100%;height:320px;border:1px solid #e5e7eb;border-radius:10px;background:#fff}.legend{display:flex;flex-wrap:wrap;gap:12px;margin-top:12px}.legend-item{display:flex;align-items:center;gap:8px;font-size:12px}.legend-color{width:12px;height:12px;border-radius:999px}
@media (max-width:960px){.grid,.grid-3{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrapper">
<div class="card">
<div class="badge $statusClass">$($Analysis.Status)</div>
<h1>ECG Diagnostics Core</h1>
<p>$($Analysis.SummaryPhrase)</p>
<div class="muted">Versao $script:ToolVersion</div>
<div class="grid-3">
<div class="kv"><strong>Maquina</strong>$($MachineInfo.ComputerName)  $($MachineInfo.MachineType)</div>
<div class="kv"><strong>Confianca</strong>$($Analysis.Confidence)</div>
<div class="kv"><strong>Observacao</strong>$($Timeline.ObservationMinutes) minuto(s)</div>
<div class="kv"><strong>Hipotese principal</strong>$($Analysis.PrimaryHypothesis)</div>
<div class="kv"><strong>Impacto</strong>$($Analysis.ImpactScope)</div>
<div class="kv"><strong>Proxima acao</strong>$($Analysis.RecommendedAction)</div>
</div>
</div>

<div class="card">
<h2>Decisao operacional</h2>
<div class="grid">
<div class="kv"><strong>Backend compartilhado</strong>$($decision.BackendStatus)<br>$($decision.BackendLine)</div>
<div class="kv"><strong>Estacao local</strong>$($decision.LocalStatus)<br>$($decision.LocalLine)</div>
</div>
<p><strong>Leitura operacional:</strong> $($Analysis.ProbablePerception)</p>
<p><strong>Leitura de lock:</strong> $($decision.LockLine)</p>
</div>

<div class="grid">
<div class="card"><h2>Reconciliacao BDE</h2><p><strong>Resumo:</strong> $remediationSummary</p><p><strong>Autorizacao informada:</strong> $(if ($null -ne $Remediation -and $Remediation.AuthorizationProvided) { 'Sim' } else { 'Nao' })</p><p><strong>Estado final:</strong> $(if ($null -ne $Remediation) { $Remediation.FinalState } else { 'N/D' })</p></div>
<div class="card"><h2>Acoes planejadas / bloqueios</h2><p><strong>Planejado</strong></p><ul>$plannedActionsItems</ul><p><strong>Bloqueios</strong></p><ul>$blockerItems</ul></div>
</div>

<div class="grid">
<div class="card"><h2>Evidencias principais</h2><ul>$supportItems</ul></div>
<div class="card"><h2>Descartes relevantes</h2><ul>$discardItems</ul></div>
</div>

<div class="grid">
<div class="card"><h2>Metricas-chave</h2><div class="grid">
<div class="kv"><strong>Pressao</strong>$($PassiveBenchmark.PressureLabel)</div>
<div class="kv"><strong>Severidade</strong>$($PassiveBenchmark.SeverityScore)/100</div>
<div class="kv"><strong>CPU media / pico</strong>$($PassiveBenchmark.AverageCpuPercent)% / $($PassiveBenchmark.PeakCpuPercent)%</div>
<div class="kv"><strong>ECGV6 media / pico</strong>$(if ($null -ne $PassiveBenchmark.AverageEcgProcessCpuPercent) { $PassiveBenchmark.AverageEcgProcessCpuPercent } else { 'N/D' })% / $(if ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent) { $PassiveBenchmark.PeakEcgProcessCpuPercent } else { 'N/D' })%</div>
<div class="kv"><strong>Locks min/med/pico</strong>$($PassiveBenchmark.MinimumLockFileCount) / $($PassiveBenchmark.AverageLockFileCount) / $($PassiveBenchmark.PeakLockFileCount)</div>
<div class="kv"><strong>DB ms media / p95</strong>$(if ($null -ne $PassiveBenchmark.AverageDatabaseProbeMs) { $PassiveBenchmark.AverageDatabaseProbeMs } else { 'N/D' }) / $(if ($null -ne $PassiveBenchmark.P95DatabaseProbeMs) { $PassiveBenchmark.P95DatabaseProbeMs } else { 'N/D' })</div>
<div class="kv"><strong>NetDir ms media / p95</strong>$(if ($null -ne $PassiveBenchmark.AverageNetDirProbeMs) { $PassiveBenchmark.AverageNetDirProbeMs } else { 'N/D' }) / $(if ($null -ne $PassiveBenchmark.P95NetDirProbeMs) { $PassiveBenchmark.P95NetDirProbeMs } else { 'N/D' })</div>
<div class="kv"><strong>DB / NetDir indisponivel</strong>$($PassiveBenchmark.DatabaseUnavailableSamples) / $($PassiveBenchmark.NetDirUnavailableSamples)</div>
<div class="kv"><strong>Timeout SMB</strong>$($PassiveBenchmark.SmbTimeoutSamples)</div>
<div class="kv"><strong>Responsavel CPU</strong>$([System.Net.WebUtility]::HtmlEncode($responsibleCpuText))</div>
<div class="kv"><strong>Capturas de processo</strong>$($PassiveBenchmark.SamplesWithProcessCapture)</div>
<div class="kv"><strong>Disk Queue pico</strong>$(if ($null -ne $PassiveBenchmark.PeakDiskQueueLength) { $PassiveBenchmark.PeakDiskQueueLength } else { 'N/D' })</div>
<div class="kv"><strong>Network Bytes pico</strong>$(if ($null -ne $PassiveBenchmark.PeakNetworkBytesTotalPerSec) { $PassiveBenchmark.PeakNetworkBytesTotalPerSec } else { 'N/D' })</div>
</div></div>
<div class="card"><h2>Pontos ainda inconclusivos</h2><ul>$inconclusiveItems</ul></div>
</div>

<details><summary>Detalhes tecnicos</summary><div style="margin-top:16px;">
<div class="card"><h3>Contexto e paths</h3><table><tr><th>Campo</th><th>Valor</th></tr><tr><td>Executavel oficial</td><td>$($Paths.ExePath) (acessivel: $($Paths.ExeAccessible))</td></tr><tr><td>Banco efetivo</td><td>$($Paths.DatabasePath) ($($Paths.DatabasePathSource)) - acessivel: $($Paths.DatabaseAccessible)</td></tr><tr><td>NetDir efetivo</td><td>$($Paths.NetDirPath) ($($Paths.NetDirSource)) - acessivel: $($Paths.NetDirAccessible)</td></tr><tr><td>Locks baseline nominal</td><td>$($PassiveBenchmark.NominalLockBaselineLikely)</td></tr><tr><td>Processo dominante no pico de CPU</td><td>$([System.Net.WebUtility]::HtmlEncode($responsibleCpuText))</td></tr></table></div>
<div class="card"><h3>Fontes observadas do BDE</h3><table><tr><th>Fonte</th><th>Caminho/Chave</th><th>Valor observado</th><th>Acessivel</th></tr>$bdeSourceRows</table></div>
<div class="card"><h3>Eventos relevantes da timeline</h3><table><tr><th>Hora</th><th>CPU%</th><th>ECGV6%</th><th>Locks</th><th>DB ms</th><th>NetDir ms</th><th>DiskQ</th><th>Net Bytes/s</th><th>Conexoes SMB</th><th>DB OK</th><th>NetDir OK</th><th>Timeout SMB</th><th>Processo dominante</th></tr>$relevantTimelineRows</table></div>
<div class="card"><h3>Processos capturados em alta CPU</h3><table><tr><th>Hora</th><th>Processo</th><th>PID</th><th>CPU amostra</th><th>CPU processo</th><th>Owner</th><th>ECG?</th></tr>$processCaptureTableRows</table></div>
<div class="card"><h3>Grafico temporal</h3><canvas id="timelineChart" width="1100" height="320"></canvas><div id="timelineLegend" class="legend"></div><p class="muted">Series normalizadas por escala propria.</p></div>
<div class="card"><h3>Hipoteses secundarias</h3><table><tr><th>Hipotese</th><th>Score</th><th>Observacao</th></tr>$secondaryRows</table></div>
</div></details>
</div>
<script>
(function(){
var chartData = $chartJson || {};
chartData.labels = chartData.labels || [];
chartData.series = chartData.series || [];
chartData.series = chartData.series.map(function(s){return{name:s.name,values:s.values,max:s.max,color:s.color};});
if(!chartData.series.length)return;
var canvas=document.getElementById('timelineChart'),legend=document.getElementById('timelineLegend');
if(!canvas||!canvas.getContext)return;
var ctx=canvas.getContext('2d'),w=canvas.width,h=canvas.height,padding={left:56,right:16,top:16,bottom:42},pw=w-padding.left-padding.right,ph=h-padding.top-padding.bottom,points=chartData.labels.length;
function xPos(i){if(points<=1)return padding.left;return padding.left+pw*i/(points-1);} function yPos(n){return padding.top+ph-ph*n;}
ctx.clearRect(0,0,w,h);ctx.fillStyle='#fff';ctx.fillRect(0,0,w,h);ctx.strokeStyle='#e5e7eb';ctx.lineWidth=1;
for(var i=0;i<=4;i++){var y=padding.top+ph*i/4;ctx.beginPath();ctx.moveTo(padding.left,y);ctx.lineTo(w-padding.right,y);ctx.stroke();}
ctx.beginPath();ctx.moveTo(padding.left,padding.top);ctx.lineTo(padding.left,h-padding.bottom);ctx.lineTo(w-padding.right,h-padding.bottom);ctx.strokeStyle='#94a3b8';ctx.stroke();
ctx.fillStyle='#475569';ctx.font='11px Segoe UI';ctx.fillText('Normalizado',padding.left,12);ctx.fillText('0',24,h-padding.bottom+4);ctx.fillText('1',24,padding.top+4);
chartData.series.forEach(function(s){var max=Number(s.max)||(s.name==='CPU %'?100:1);ctx.beginPath();var started=false;for(var idx=0;idx<s.values.length;idx++){var v=s.values[idx];if(v===null||v===undefined)continue;var norm=Number(v)/max;if(norm<0)norm=0;if(norm>1)norm=1;var x=xPos(idx),y=yPos(norm);if(!started){ctx.moveTo(x,y);started=true;}else{ctx.lineTo(x,y);}}ctx.strokeStyle=s.color;ctx.lineWidth=2;ctx.stroke();});
chartData.labels.forEach(function(l,i){if(i%3===0||i===chartData.labels.length-1){var x=xPos(i);ctx.save();ctx.translate(x,h-padding.bottom+14);ctx.rotate(-0.45);ctx.fillStyle='#475569';ctx.font='10px Segoe UI';ctx.fillText(l,0,0);ctx.restore();}});
chartData.series.forEach(function(s){var leg=document.createElement('div');leg.className='legend-item';var cb=document.createElement('span');cb.className='legend-color';cb.style.backgroundColor=s.color;var txt=document.createElement('span');txt.textContent=s.name+' (max='+s.max+')';leg.appendChild(cb);leg.appendChild(txt);legend.appendChild(leg);});
})();
</script>
</body>
</html>
"@
}

function Build-JsonReport {
    param($Analysis, $MachineInfo, $Paths, $Timeline, $PassiveBenchmark, $Profile, [string]$ModeName, [string[]]$AppliedChanges = @(), [string]$RollbackFile = '', $Remediation = $null)
    $stationRole = Get-ProfileString -Profile $Profile -Name 'StationRole' -Default 'AUTO'
    $stationAlias = Get-ProfileString -Profile $Profile -Name 'StationAlias' -Default ''
    $decision = Get-OperationalDecisionModel -Analysis $Analysis -Paths $Paths -PassiveBenchmark $PassiveBenchmark
    return [PSCustomObject]@{
        Metadata = [PSCustomObject]@{
            ToolName = $script:ToolName
            ToolVersion = $script:ToolVersion
            RunId = $script:RunId
            Mode = $ModeName
            GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            ComputerName = $MachineInfo.ComputerName
            MachineType = $MachineInfo.MachineType
            ExecutedBy = $MachineInfo.ExecutedBy
            StationRole = $stationRole
            StationAlias = $stationAlias
            CoreScriptPath = $script:ScriptPath
            ProfilePath = $ProfilePath
        }
        Context = [PSCustomObject]@{
            DesiredDbPath = $Paths.DesiredDbPath
            EffectiveDbPath = $Paths.DatabasePath
            DesiredNetDir = $Paths.DesiredNetDir
            EffectiveNetDir = $Paths.NetDirPath
            DesiredExePath = $Paths.DesiredExePath
            EffectiveExePath = $Paths.ExePath
            DatabaseAccessible = $Paths.DatabaseAccessible
            NetDirAccessible = $Paths.NetDirAccessible
            ExeAccessible = $Paths.ExeAccessible
            DatabasePathSource = $Paths.DatabasePathSource
            NetDirSource = $Paths.NetDirSource
        }
        OperationalDecision = $decision
        PassiveBenchmark = $PassiveBenchmark
        Analysis = $Analysis
        Timeline = $Timeline
        BdeSources = $Paths.BdeSources
        Remediation = $Remediation
        AppliedChanges = @($AppliedChanges)
        RollbackFile = $RollbackFile
        Logs = @($script:LogLines)
    }
}

function Get-CompareSourceFiles {
    param([string]$RootPath)
    if (-not (Test-Path $RootPath)) { throw "OutDir nao encontrado: $RootPath" }
    $files = Get-ChildItem -Path $RootPath -Recurse -Filter 'ECG_Report.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($files.Count -lt 2) { throw "Preciso de pelo menos 2 relatorios JSON em $RootPath" }
    return @($files[0].FullName, $files[1].FullName)
}

function Normalize-CompareValue {
    param([string]$Value)
    return Normalize-PathString $Value
}

function Build-CompareHtml {
    param($LeftData, $RightData, [string]$LeftPath, [string]$RightPath, $Rows, [string]$Status)
    $rowHtml = ($Rows | ForEach-Object {
        $css = if ($_.Match) { 'ok' } else { 'drift' }
        "<tr class='$css'><td>$([System.Net.WebUtility]::HtmlEncode($_.Field))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Left))</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Right))</td><td>$([System.Net.WebUtility]::HtmlEncode($_.Result))</td></tr>"
    }) -join ''
    $statusCss = if ($Status -eq 'CONVERGENTE') { 'ok' } else { 'drift' }

    return @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>ECG Compare Report</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;background:#f3f6fb;color:#1f2937;margin:0}
.wrapper{max-width:1100px;margin:0 auto;padding:24px}
.card{background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(15,23,42,.08);padding:20px;margin-bottom:16px}
.badge{display:inline-block;padding:8px 14px;border-radius:999px;font-weight:700;font-size:13px}
.badge.ok{background:#dcfce7;color:#166534}.badge.drift{background:#fee2e2;color:#991b1b}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #e5e7eb;padding:8px 10px;text-align:left;vertical-align:top}th{background:#f9fafb}tr.ok{background:#f0fdf4}tr.drift{background:#fef2f2}.muted{color:#6b7280;font-size:12px}
</style>
</head>
<body>
<div class="wrapper">
<div class="card">
<div class="badge $statusCss">$Status</div>
<h1>ECG Diagnostics Compare</h1>
<p>Comparacao de convergencia entre dois laudos JSON.</p>
<p class="muted"><strong>Esquerda:</strong> $([System.Net.WebUtility]::HtmlEncode($LeftPath))<br><strong>Direita:</strong> $([System.Net.WebUtility]::HtmlEncode($RightPath))</p>
</div>
<div class="card">
<h2>Campos criticos</h2>
<table>
<tr><th>Campo</th><th>Esquerda</th><th>Direita</th><th>Status</th></tr>
$rowHtml
</table>
</div>
</div>
</body>
</html>
"@
}

function Invoke-FixAutoMode {
    param([bool]$ApplyFixes)
    $modeLabel = if ($Mode -eq 'Detect') { 'Detect' } elseif ($ApplyFixes) { 'Fix' } else { 'Auto' }
    Log "Modo $modeLabel - Iniciado" 'STEP'

    $profile = Read-ProfileIni -Path $ProfilePath
    $effectiveOutDir = Resolve-EffectiveOutDir -Profile $profile
    $expectedDb = Normalize-PathString (Get-ProfileString -Profile $profile -Name 'ExpectedDbPath' -Default '\\192.168.1.57\Database')
    $expectedNetDir = Normalize-PathString (Get-ProfileString -Profile $profile -Name 'ExpectedNetDir' -Default (Join-Path $expectedDb 'NetDir'))
    $observeMinutes = [Math]::Max(1, (Get-ProfileInt -Profile $profile -Name 'MonitorMinutes' -Default 3))
    $sampleInterval = [Math]::Max(5, (Get-ProfileInt -Profile $profile -Name 'SampleIntervalSeconds' -Default 15))
    $processCaptureThreshold = [Math]::Max(60, (Get-ProfileInt -Profile $profile -Name 'CpuProcessCaptureThreshold' -Default 80))
    $topProcessCount = [Math]::Min(5, [Math]::Max(1, (Get-ProfileInt -Profile $profile -Name 'TopProcessCaptureCount' -Default 3)))
    $enableLatencyMetrics = Get-ProfileBool -Profile $profile -Name 'EnableLatencyMetrics' -Default $true
    $enableEcgProcessMetrics = Get-ProfileBool -Profile $profile -Name 'EnableEcgProcessMetrics' -Default $true
    $enableDiskMetrics = Get-ProfileBool -Profile $profile -Name 'EnableDiskMetrics' -Default $false
    $enableNetworkMetrics = Get-ProfileBool -Profile $profile -Name 'EnableNetworkMetrics' -Default $false
    $setHwPath = Get-ProfileBool -Profile $profile -Name 'SetMachineHwPath' -Default $false

    $prePaths = Resolve-EcgPaths -Profile $profile
    $preSourceSnapshot = $prePaths.BdeSources
    $remediationPlan = Build-RemediationPlan -SourceSnapshot $preSourceSnapshot -ApplyFixes:$ApplyFixes -AuthorizationProvided:$AuthorizedRemediation
    foreach ($plannedChange in @($remediationPlan.PlannedActions)) { Log ("PLANO: {0}" -f $plannedChange) 'STEP' }
    foreach ($blocker in @($remediationPlan.Blockers)) { Log ("BLOQUEIO: {0}" -f $blocker) 'WARN' }

    if ($ApplyFixes -and -not $AuthorizedRemediation) {
        throw 'Modo Fix requer autorizacao explicita via -AuthorizedRemediation.'
    }

    if ($ApplyFixes -and -not (Test-IsAdmin)) {
        throw 'Modo Fix requer execucao elevada (Administrador) para alterar HKLM/IDAPI32.CFG com seguranca.'
    }

    $machineInfo = Get-KnownMachineInfo -ComputerName $script:HostName -Profile $profile
    $runRoot = Join-Path $effectiveOutDir $script:RunId
    $script:CurrentRunRoot = $runRoot
    New-Item -ItemType Directory -Path $runRoot -Force | Out-Null

    $appliedChanges = @()
    $rollbackFile = ''

    if ($ApplyFixes -and $remediationPlan.Blockers.Count -gt 0) {
        throw ('Remediacao bloqueada: ' + ($remediationPlan.Blockers -join ' | '))
    }

    if ($ApplyFixes) {
        $rollbackFile = Join-Path $runRoot 'BDE_Rollback.reg'
        Export-BdeRollbackFile -Path $rollbackFile | Out-Null
        $appliedChanges += "Backup de rollback do registro gerado em $rollbackFile"

        $regPaths = @(
            'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT',
            'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT',
            'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT'
        )

        foreach ($rp in $regPaths) {
            $current = Normalize-PathString (Get-RegistryValueSafe -Path $rp -Name 'NETDIR')
            if ($current -ne $expectedNetDir) {
                if (Set-RegistryValueSafe -Path $rp -Name 'NETDIR' -Value $expectedNetDir) {
                    $appliedChanges += "NETDIR ajustado em $rp para $expectedNetDir"
                }
                else {
                    throw "Falha ao ajustar NETDIR em $rp"
                }
            }
            else {
                $appliedChanges += "NETDIR ja estava aderente em $rp"
            }
        }

        if ($setHwPath) {
            $hwChanges = @(Set-HwCaminhoDbSafe -Value $expectedDb)
            if ($hwChanges.Count -gt 0) { $appliedChanges += $hwChanges }
        }

        $idapiPaths = @(
            'C:\Program Files (x86)\Common Files\Borland Shared\BDE\IDAPI32.CFG',
            'C:\Program Files\Common Files\Borland Shared\BDE\IDAPI32.CFG'
        )

        foreach ($idapi in $idapiPaths) {
            if (-not (Test-Path $idapi)) { continue }
            $content = [System.IO.File]::ReadAllText($idapi, [System.Text.Encoding]::Default)
            $updated = $content

            if ($content -match '(?im)^\s*NET\s+DIR\s*=\s*(.*)$') {
                $old = Normalize-PathString $matches[1].Trim()
                if ($old -ne $expectedNetDir) {
                    $updated = [regex]::Replace($content, '(?im)^\s*NET\s+DIR\s*=.*$', "NET DIR = $expectedNetDir")
                }
            }
            elseif ($content -match '(?im)^\s*NETDIR\s*=\s*(.*)$') {
                $old = Normalize-PathString $matches[1].Trim()
                if ($old -ne $expectedNetDir) {
                    $updated = [regex]::Replace($content, '(?im)^\s*NETDIR\s*=.*$', "NETDIR = $expectedNetDir")
                }
            }
            else {
                $updated = $content + [Environment]::NewLine + "NET DIR = $expectedNetDir" + [Environment]::NewLine
            }

            if ($updated -ne $content) {
                Copy-Item -Path $idapi -Destination "$idapi.bak" -Force
                [System.IO.File]::WriteAllText($idapi, $updated, [System.Text.Encoding]::Default)
                $appliedChanges += "IDAPI32.CFG corrigido (backup em $idapi.bak)"
            }
            else {
                $appliedChanges += "IDAPI32.CFG ja aderente em $idapi"
            }
        }
    }

    $paths = Resolve-EcgPaths -Profile $profile
    $postSourceSnapshot = $paths.BdeSources
    $remediationPlan = Build-RemediationPlan -SourceSnapshot $postSourceSnapshot -ApplyFixes:$ApplyFixes -AuthorizationProvided:$AuthorizedRemediation -AppliedChanges $appliedChanges
    $timeline = Collect-Timeline -MachineType $machineInfo.MachineType -Paths $paths -Minutes $observeMinutes -IntervalSeconds $sampleInterval -ProcessCaptureThresholdPercent $processCaptureThreshold -TopProcessCount $topProcessCount -EnableLatencyMetrics $enableLatencyMetrics -EnableEcgProcessMetrics $enableEcgProcessMetrics -EnableDiskMetrics $enableDiskMetrics -EnableNetworkMetrics $enableNetworkMetrics
    $benchmark = Build-PassiveBenchmark -Timeline $timeline -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'
    $analysis = Build-AnalysisModel -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'

    $html = Build-HtmlReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO' -Remediation $remediationPlan
    $jsonData = Build-JsonReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -Profile $profile -ModeName $modeLabel -AppliedChanges $appliedChanges -RollbackFile $rollbackFile -Remediation $remediationPlan

    $htmlPath = Join-Path $runRoot 'ECG_Report.html'
    $jsonPath = Join-Path $runRoot 'ECG_Report.json'
    Write-Utf8File -Path $htmlPath -Text $html
    Write-Utf8File -Path $jsonPath -Text ($jsonData | ConvertTo-Json -Depth 10)

    if ($OpenReport) { Start-Process $htmlPath }
    Log "Relatorio HTML salvo em $htmlPath" 'STEP'
    Log "Relatorio JSON salvo em $jsonPath" 'STEP'
}

function Invoke-CompareMode {
    $profile = Read-ProfileIni -Path $ProfilePath
    $effectiveOutDir = Resolve-EffectiveOutDir -Profile $profile
    $script:CurrentRunRoot = $effectiveOutDir

    $left = ''
    $right = ''

    if (-not [string]::IsNullOrWhiteSpace($CompareLeftReport) -and -not [string]::IsNullOrWhiteSpace($CompareRightReport)) {
        $left = $CompareLeftReport
        $right = $CompareRightReport
    }
    else {
        $pair = Get-CompareSourceFiles -RootPath $effectiveOutDir
        $left = $pair[0]
        $right = $pair[1]
    }

    if (-not (Test-Path $left)) { throw "Relatorio esquerdo nao encontrado: $left" }
    if (-not (Test-Path $right)) { throw "Relatorio direito nao encontrado: $right" }

    Log "Comparando $left com $right" 'STEP'
    $leftData = Get-Content -Path $left -Raw | ConvertFrom-Json
    $rightData = Get-Content -Path $right -Raw | ConvertFrom-Json

    $rows = @(
        [PSCustomObject]@{ Field = 'DesiredDbPath'; Left = $leftData.Context.DesiredDbPath; Right = $rightData.Context.DesiredDbPath },
        [PSCustomObject]@{ Field = 'EffectiveDbPath'; Left = $leftData.Context.EffectiveDbPath; Right = $rightData.Context.EffectiveDbPath },
        [PSCustomObject]@{ Field = 'DesiredNetDir'; Left = $leftData.Context.DesiredNetDir; Right = $rightData.Context.DesiredNetDir },
        [PSCustomObject]@{ Field = 'EffectiveNetDir'; Left = $leftData.Context.EffectiveNetDir; Right = $rightData.Context.EffectiveNetDir }
    ) | ForEach-Object {
        $leftNorm = Normalize-CompareValue -Value ([string]$_.Left)
        $rightNorm = Normalize-CompareValue -Value ([string]$_.Right)
        $match = ($leftNorm -eq $rightNorm)
        $resultLabel = if ($match) { 'OK' } else { 'DRIFT' }
        [PSCustomObject]@{ Field = $_.Field; Left = $_.Left; Right = $_.Right; Match = $match; Result = $resultLabel }
    }

    $status = if (@($rows | Where-Object { -not $_.Match }).Count -eq 0) { 'CONVERGENTE' } else { 'DRIFT_CRITICO' }
    $html = Build-CompareHtml -LeftData $leftData -RightData $rightData -LeftPath $left -RightPath $right -Rows $rows -Status $status
    $outHtml = Join-Path $effectiveOutDir ("Compare_{0}.html" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
    Write-Utf8File -Path $outHtml -Text $html

    if ($OpenReport) { Start-Process $outHtml }
    Log "Comparacao concluida: $status" 'STEP'
    Log "Relatorio de comparacao salvo em $outHtml" 'STEP'
}

function Invoke-RollbackMode {
    if (-not $RollbackFile) { throw 'Parametro -RollbackFile obrigatorio' }
    if (-not (Test-Path $RollbackFile)) { throw "Arquivo $RollbackFile nao encontrado" }
    if (-not (Test-IsAdmin)) { throw 'Rollback requer privilegios administrativos' }
    $script:CurrentRunRoot = Split-Path -Parent $RollbackFile
    $process = Start-Process -FilePath 'reg.exe' -ArgumentList @('import', $RollbackFile) -Wait -NoNewWindow -PassThru
    if ($process.ExitCode -ne 0) { throw "reg import retornou codigo $($process.ExitCode)" }
    Log "Rollback executado a partir de $RollbackFile" 'STEP'
}

function Invoke-MonitorMode {
    Log 'Modo Monitor - Coleta prolongada (leitura apenas)' 'STEP'
    $profile = Read-ProfileIni -Path $ProfilePath
    $effectiveOutDir = Resolve-EffectiveOutDir -Profile $profile
    $monitorMinutes = [Math]::Max(1, (Get-ProfileInt -Profile $profile -Name 'MonitorMinutes' -Default 10))
    $sampleInterval = [Math]::Max(5, (Get-ProfileInt -Profile $profile -Name 'SampleIntervalSeconds' -Default 20))
    $processCaptureThreshold = [Math]::Max(60, (Get-ProfileInt -Profile $profile -Name 'CpuProcessCaptureThreshold' -Default 80))
    $topProcessCount = [Math]::Min(5, [Math]::Max(1, (Get-ProfileInt -Profile $profile -Name 'TopProcessCaptureCount' -Default 3)))
    $enableLatencyMetrics = Get-ProfileBool -Profile $profile -Name 'EnableLatencyMetrics' -Default $true
    $enableEcgProcessMetrics = Get-ProfileBool -Profile $profile -Name 'EnableEcgProcessMetrics' -Default $true
    $enableDiskMetrics = Get-ProfileBool -Profile $profile -Name 'EnableDiskMetrics' -Default $false
    $enableNetworkMetrics = Get-ProfileBool -Profile $profile -Name 'EnableNetworkMetrics' -Default $false

    $runRoot = Join-Path $effectiveOutDir $script:RunId
    $script:CurrentRunRoot = $runRoot
    New-Item -ItemType Directory -Path $runRoot -Force | Out-Null

    $paths = Resolve-EcgPaths -Profile $profile
    $machineInfo = Get-KnownMachineInfo -ComputerName $script:HostName -Profile $profile
    $timeline = Collect-Timeline -MachineType $machineInfo.MachineType -Paths $paths -Minutes $monitorMinutes -IntervalSeconds $sampleInterval -ProcessCaptureThresholdPercent $processCaptureThreshold -TopProcessCount $topProcessCount -EnableLatencyMetrics $enableLatencyMetrics -EnableEcgProcessMetrics $enableEcgProcessMetrics -EnableDiskMetrics $enableDiskMetrics -EnableNetworkMetrics $enableNetworkMetrics
    $benchmark = Build-PassiveBenchmark -Timeline $timeline -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'
    $analysis = Build-AnalysisModel -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'

    $html = Build-HtmlReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'
    $jsonData = Build-JsonReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -Profile $profile -ModeName 'Monitor'

    $htmlPath = Join-Path $runRoot 'ECG_Report.html'
    $jsonPath = Join-Path $runRoot 'ECG_Report.json'
    Write-Utf8File -Path $htmlPath -Text $html
    Write-Utf8File -Path $jsonPath -Text ($jsonData | ConvertTo-Json -Depth 10)

    if ($OpenReport) { Start-Process $htmlPath }
    Log "Relatorio de monitoramento salvo em $htmlPath" 'STEP'
    Log "Relatorio JSON de monitoramento salvo em $jsonPath" 'STEP'
}

function Invoke-CollectStaticMode {
    Log 'Coletando informacoes estaticas' 'STEP'
    $profile = Read-ProfileIni -Path $ProfilePath
    $effectiveOutDir = Resolve-EffectiveOutDir -Profile $profile
    $outDir = Join-Path $effectiveOutDir 'static'
    $script:CurrentRunRoot = $outDir
    if (-not (Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $hostName = $env:COMPUTERNAME
    $user = try { [Security.Principal.WindowsIdentity]::GetCurrent().Name } catch { $env:USERNAME }
    $shares = Get-LocalShares
    $netConns = Get-NetworkConnections
    $envVars = Get-ChildItem Env: | Sort-Object Name | ForEach-Object { [PSCustomObject]@{ Name = $_.Name; Value = $_.Value } }
    $regValues = [PSCustomObject]@{
        HKLM = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT' -Name 'NETDIR'
        HKLM_WOW6432 = Get-RegistryValueSafe -Path 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT' -Name 'NETDIR'
        HKCU = Get-RegistryValueSafe -Path 'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT' -Name 'NETDIR'
    }

    $summary = [PSCustomObject]@{
        Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Host = $hostName
        User = $user
        IsAdmin = (Test-IsAdmin)
        ProfilePath = $ProfilePath
        CoreScriptPath = $script:ScriptPath
        ExpectedDbPath = Get-ProfileString -Profile $profile -Name 'ExpectedDbPath' -Default '\\192.168.1.57\Database'
        ExpectedNetDir = Get-ProfileString -Profile $profile -Name 'ExpectedNetDir' -Default '\\192.168.1.57\Database\NetDir'
        HW_CAMINHO_DB_Process = $env:HW_CAMINHO_DB
        NETDIR_Registry = $regValues
        Shares = $shares
        NetworkConnections = $netConns
        EnvironmentVariables = $envVars
    }

    $jsonFile = Join-Path $outDir ("ECG_State_{0}_{1}.json" -f $hostName, $stamp)
    Write-Utf8File -Path $jsonFile -Text ($summary | ConvertTo-Json -Depth 8)
    Log "Coleta estatica concluida: $jsonFile" 'STEP'
}



function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}





function Test-IsIpAddress {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $false }
    $nullIp = $null
    return [System.Net.IPAddress]::TryParse($Value, [ref]$nullIp)
}


function Get-UncHost {
    param([string]$Path)
    $p = Normalize-PathString $Path
    if ($p -and $p -match '^\\\\([^\\]+)\\') { return [string]$matches[1] }
    return $null
}


function Get-TextOrDefault {
    param(
        [object]$Value,
        [string]$Default = 'N/D',
        [string]$Suffix = ''
    )
    if ($null -eq $Value) { return $Default }
    $text = [string]$Value
    if ([string]::IsNullOrWhiteSpace($text)) { return $Default }
    return ($text + $Suffix)
}

function Get-Stats {
    param([object[]]$Values)

    $nums = New-Object System.Collections.Generic.List[double]
    foreach ($v in @($Values)) {
        $d = ConvertTo-DoubleSafe -Value $v
        if ($null -ne $d) { [void]$nums.Add([double]$d) }
    }

    if ($nums.Count -eq 0) {
        return [PSCustomObject]@{
            Count = 0
            Min = $null
            Avg = $null
            P95 = $null
            Max = $null
        }
    }

    $ordered = @($nums | Sort-Object)
    $sum = 0.0
    foreach ($n in $ordered) { $sum += $n }
    $idx = [Math]::Ceiling(0.95 * $ordered.Count) - 1
    if ($idx -lt 0) { $idx = 0 }
    if ($idx -ge $ordered.Count) { $idx = $ordered.Count - 1 }

    return [PSCustomObject]@{
        Count = $ordered.Count
        Min = [math]::Round($ordered[0], 2)
        Avg = [math]::Round(($sum / $ordered.Count), 2)
        P95 = [math]::Round($ordered[$idx], 2)
        Max = [math]::Round($ordered[-1], 2)
    }
}

function Read-IniFile {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { throw "INI nao encontrado: $Path" }

    $ini = @{}
    $section = 'General'
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

function Get-IniInt {
    param($Ini, [string]$Section, [string]$Key, [int]$Default)
    $raw = Get-IniString -Ini $Ini -Section $Section -Key $Key -Default ''
    $tmp = 0
    if ([int]::TryParse($raw, [ref]$tmp)) { return $tmp }
    return $Default
}

function Get-IniBool {
    param($Ini, [string]$Section, [string]$Key, [bool]$Default = $false)
    $raw = (Get-IniString -Ini $Ini -Section $Section -Key $Key -Default '').ToLowerInvariant()
    if ([string]::IsNullOrWhiteSpace($raw)) { return $Default }
    if (@('1','true','yes','sim','y','on') -contains $raw) { return $true }
    if (@('0','false','no','nao','nao','off') -contains $raw) { return $false }
    return $Default
}

function Resolve-CompareOutDir {
    param($Ini)
    if ($script:CliOutDirProvided) { return $OutDir }
    $profileOut = Get-IniString -Ini $Ini -Section 'General' -Key 'OutDir' -Default ''
    if (-not [string]::IsNullOrWhiteSpace($profileOut)) { return $profileOut }
    return 'C:\ECG\FieldKit\out'
}

function Resolve-HostInfo {
    param([string]$HostName)

    $resolved = New-Object System.Collections.Generic.List[string]
    $dnsOk = $false
    $dnsError = $null

    if ([string]::IsNullOrWhiteSpace($HostName)) {
        return [PSCustomObject]@{
            InputHost = $HostName
            IsIp = $false
            ResolvedIPs = @()
            DnsOk = $false
            DnsError = 'Host vazio'
        }
    }

    if (Test-IsIpAddress $HostName) {
        return [PSCustomObject]@{
            InputHost = $HostName
            IsIp = $true
            ResolvedIPs = @($HostName)
            DnsOk = $true
            DnsError = $null
        }
    }

    try {
        if (Get-Command Resolve-DnsName -ErrorAction SilentlyContinue) {
            $answers = @(Resolve-DnsName -Name $HostName -Type A -ErrorAction Stop)
            foreach ($a in $answers) {
                if ($a.IPAddress -and (-not $resolved.Contains([string]$a.IPAddress))) { [void]$resolved.Add([string]$a.IPAddress) }
            }
        }
        else {
            $addresses = [System.Net.Dns]::GetHostAddresses($HostName)
            foreach ($a in $addresses) {
                $ip = [string]$a.IPAddressToString
                if (-not [string]::IsNullOrWhiteSpace($ip) -and (-not $resolved.Contains($ip))) { [void]$resolved.Add($ip) }
            }
        }
        $dnsOk = $resolved.Count -gt 0
    }
    catch {
        $dnsError = $_.Exception.Message
    }

    return [PSCustomObject]@{
        InputHost = $HostName
        IsIp = $false
        ResolvedIPs = @($resolved)
        DnsOk = $dnsOk
        DnsError = $dnsError
    }
}

function Test-HostPing {
    param([string]$HostName, [int]$Count)

    $times = New-Object System.Collections.Generic.List[double]
    $errorText = $null

    try {
        $results = @(Test-Connection -ComputerName $HostName -Count $Count -ErrorAction Stop)
        foreach ($r in $results) {
            if ($null -ne $r.ResponseTime) { [void]$times.Add([double]$r.ResponseTime) }
        }
    }
    catch {
        $errorText = $_.Exception.Message
    }

    $stats = Get-Stats -Values @($times)
    return [PSCustomObject]@{
        Host = $HostName
        Sent = $Count
        Received = $times.Count
        Lost = ($Count - $times.Count)
        MinMs = $stats.Min
        AvgMs = $stats.Avg
        P95Ms = $stats.P95
        MaxMs = $stats.Max
        Error = $errorText
    }
}

function Test-TcpPortLatency {
    param([string]$HostName, [int]$Port, [int]$TimeoutMs)

    $client = New-Object System.Net.Sockets.TcpClient
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        $iar = $client.BeginConnect($HostName, $Port, $null, $null)
        if (-not $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            throw "Timeout ao conectar em $HostName`:$Port"
        }
        $client.EndConnect($iar)
        $sw.Stop()
        return [PSCustomObject]@{
            Host = $HostName
            Port = $Port
            Reachable = $true
            ConnectMs = [math]::Round([double]$sw.Elapsed.TotalMilliseconds, 2)
            Error = $null
        }
    }
    catch {
        $sw.Stop()
        return [PSCustomObject]@{
            Host = $HostName
            Port = $Port
            Reachable = $false
            ConnectMs = [math]::Round([double]$sw.Elapsed.TotalMilliseconds, 2)
            Error = $_.Exception.Message
        }
    }
    finally {
        $client.Close()
    }
}

function Measure-PathProbe {
    param([string]$Path)

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $exists = $false
    $errorText = $null
    try {
        $normalized = Normalize-PathString $Path
        $exists = Test-Path -LiteralPath $normalized
        if ($exists) { Get-Item -LiteralPath $normalized -ErrorAction Stop | Out-Null }
    }
    catch {
        $exists = $false
        $errorText = $_.Exception.Message
    }
    finally {
        $sw.Stop()
    }

    return [PSCustomObject]@{
        Accessible = $exists
        ProbeMs = [math]::Round([double]$sw.Elapsed.TotalMilliseconds, 2)
        Error = $errorText
    }
}


function Get-SmbSnapshot {
    param([string[]]$Hosts)

    $rows = @()
    $normalizedHosts = @()
    foreach ($h in @($Hosts)) {
        if ([string]::IsNullOrWhiteSpace([string]$h)) { continue }
        $upper = ([string]$h).ToUpperInvariant()
        if ($normalizedHosts -notcontains $upper) { $normalizedHosts += $upper }
        if ($upper -match '^([^\.]+)\.') {
            $short = [string]$matches[1]
            if ($normalizedHosts -notcontains $short) { $normalizedHosts += $short }
        }
    }

    try {
        if (Get-Command Get-SmbConnection -ErrorAction SilentlyContinue) {
            $all = @(Get-SmbConnection -ErrorAction SilentlyContinue)
            foreach ($conn in $all) {
                $server = ''
                if ($null -ne $conn.ServerName) { $server = ([string]$conn.ServerName).ToUpperInvariant() }
                $serverShort = $server
                if ($server -match '^([^\.]+)\.') { $serverShort = [string]$matches[1] }
                if ($normalizedHosts.Count -eq 0 -or $normalizedHosts -contains $server -or $normalizedHosts -contains $serverShort) {
                    $rows += [PSCustomObject]@{
                        ServerName = [string]$conn.ServerName
                        ShareName = [string]$conn.ShareName
                        UserName = [string]$conn.UserName
                        NumOpens = $conn.NumOpens
                        Dialect = [string]$conn.Dialect
                    }
                }
            }
        }
    }
    catch {}

    return $rows
}

function Get-NetworkSnapshot {
    $rows = @()
    try {
        if (Get-Command Get-NetIPConfiguration -ErrorAction SilentlyContinue) {
            $items = @(Get-NetIPConfiguration | Where-Object { $_.IPv4Address })
            foreach ($i in $items) {
                $gateway = $null
                if ($i.IPv4DefaultGateway) { $gateway = $i.IPv4DefaultGateway.NextHop }
                $rows += [PSCustomObject]@{
                    InterfaceAlias = $i.InterfaceAlias
                    IPv4 = @($i.IPv4Address | ForEach-Object { $_.IPAddress }) -join ', '
                    Gateway = $gateway
                    DnsServer = @($i.DNSServer.ServerAddresses) -join ', '
                }
            }
        }
    }
    catch {}
    return $rows
}

function ConvertTo-HtmlTable {
    param([object[]]$Rows)

    if (-not $Rows -or @($Rows).Count -eq 0) {
        return '<p class="muted">Sem dados.</p>'
    }

    $props = $Rows[0].PSObject.Properties.Name
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<table><tr>')
    foreach ($p in $props) {
        [void]$sb.Append('<th>' + [System.Net.WebUtility]::HtmlEncode($p) + '</th>')
    }
    [void]$sb.Append('</tr>')

    foreach ($row in $Rows) {
        [void]$sb.Append('<tr>')
        foreach ($p in $props) {
            $val = $row.$p
            if ($val -is [System.Collections.IEnumerable] -and -not ($val -is [string])) {
                $val = (@($val) -join ', ')
            }
            [void]$sb.Append('<td>' + [System.Net.WebUtility]::HtmlEncode([string]$val) + '</td>')
        }
        [void]$sb.Append('</tr>')
    }
    [void]$sb.Append('</table>')
    return $sb.ToString()
}

function Get-EnabledTargets {
    param($Ini)

    $targets = @()
    foreach ($section in @($Ini.Keys | Sort-Object)) {
        if ($section -notmatch '^Target\d+$') { continue }
        $enabled = Get-IniBool -Ini $Ini -Section $section -Key 'Enabled' -Default $false
        if (-not $enabled) { continue }

        $label = Get-IniString -Ini $Ini -Section $section -Key 'Label' -Default $section
        $dbPath = Normalize-PathString (Get-IniString -Ini $Ini -Section $section -Key 'DbPath' -Default '')
        $netDirPath = Normalize-PathString (Get-IniString -Ini $Ini -Section $section -Key 'NetDirPath' -Default '')
        $hostOsHint = Get-IniString -Ini $Ini -Section $section -Key 'HostOsHint' -Default ''
        $smbDialectHint = Get-IniString -Ini $Ini -Section $section -Key 'SmbDialectHint' -Default ''
        $legacyHint = Get-IniBool -Ini $Ini -Section $section -Key 'LegacyHint' -Default $false
        $notes = Get-IniString -Ini $Ini -Section $section -Key 'Notes' -Default ''

        if (-not (Test-IsUncPath -Path $dbPath)) { continue }
        if (-not (Test-IsUncPath -Path $netDirPath)) { continue }

        $targets += [PSCustomObject]@{
            Section = $section
            Label = $label
            DbPath = $dbPath
            NetDirPath = $netDirPath
            DbHost = Get-UncHost -Path $dbPath
            NetDirHost = Get-UncHost -Path $netDirPath
            HostOsHint = $hostOsHint
            SmbDialectHint = $smbDialectHint
            LegacyHint = $legacyHint
            Notes = $notes
        }
    }

    return $targets
}

function Build-TargetAssessment {
    param(
        $DbPing,
        $NetPing,
        $DbTcp,
        $NetTcp,
        $DbStats,
        $NetStats,
        $LocalStats,
        [int]$DbUnavailable,
        [int]$NetUnavailable,
        [double]$AvgCpu,
        [double]$PeakCpu
    )

    $findings = New-Object System.Collections.Generic.List[string]
    $actions = New-Object System.Collections.Generic.List[string]
    $classification = 'INCONCLUSIVO'
    $confidence = 'Media'

    $dbPingPartialLoss = ($DbPing.Received -gt 0 -and $DbPing.Lost -gt 0)
    $netPingPartialLoss = ($NetPing.Received -gt 0 -and $NetPing.Lost -gt 0)
    $networkTrouble = $false

    if ((-not $DbTcp.Reachable) -or (-not $NetTcp.Reachable)) {
        $networkTrouble = $true
    }
    elseif (($dbPingPartialLoss -or $netPingPartialLoss) -and ((($DbPing.AvgMs -ne $null) -and $DbPing.AvgMs -ge 50) -or (($NetPing.AvgMs -ne $null) -and $NetPing.AvgMs -ge 50))) {
        $networkTrouble = $true
    }

    $remoteAvgValues = New-Object System.Collections.Generic.List[double]
    if ($null -ne $DbStats.Avg) { [void]$remoteAvgValues.Add([double]$DbStats.Avg) }
    if ($null -ne $NetStats.Avg) { [void]$remoteAvgValues.Add([double]$NetStats.Avg) }
    $remoteAvg = $null
    if ($remoteAvgValues.Count -gt 0) {
        $remoteAvg = [math]::Round((($remoteAvgValues | Measure-Object -Average).Average), 2)
    }

    $remoteToLocalRatio = $null
    if ($null -ne $remoteAvg -and $null -ne $LocalStats.Avg -and [double]$LocalStats.Avg -gt 0) {
        $remoteToLocalRatio = [math]::Round(($remoteAvg / [double]$LocalStats.Avg), 2)
    }

    $shareSlow = $false
    if (($null -ne $DbStats.P95 -and $DbStats.P95 -ge 100) -or ($null -ne $NetStats.P95 -and $NetStats.P95 -ge 100)) {
        $shareSlow = $true
    }
    elseif ($DbUnavailable -gt 0 -or $NetUnavailable -gt 0) {
        $shareSlow = $true
    }
    elseif (($null -ne $remoteAvg -and $remoteAvg -ge 20) -and ($null -ne $remoteToLocalRatio -and $remoteToLocalRatio -ge 10)) {
        $shareSlow = $true
    }

    $localPressure = $false
    if (($null -ne $LocalStats.P95 -and $LocalStats.P95 -ge 15) -or ($null -ne $AvgCpu -and $AvgCpu -ge 70) -or ($null -ne $PeakCpu -and $PeakCpu -ge 90)) {
        $localPressure = $true
    }

    if ($networkTrouble) {
        $classification = 'REDE/CAMINHO'
        $confidence = 'Alta'
        [void]$findings.Add('Falha objetiva de transporte: TCP 445 indisponivel e/ou perda parcial relevante com latencia alta.')
        [void]$actions.Add('Validar cabo, switch, rota, firewall e reachability ate o host do share.')
    }
    elseif ($shareSlow -and -not $localPressure) {
        $classification = 'SERVIDOR/SHARE'
        $confidence = 'Media'
        [void]$findings.Add('Acesso UNC mais lento que o baseline local, sem pressao local equivalente.')
        [void]$actions.Add('Validar resposta do servidor, storage, antivirus em tempo real e compatibilidade do share.')
    }
    elseif ($localPressure -and -not $shareSlow) {
        $classification = 'ESTACAO LOCAL'
        $confidence = 'Baixa'
        [void]$findings.Add('Baseline local pressionado sem evidencia equivalente de degradacao no share.')
        [void]$actions.Add('Validar CPU, disco, antivirus e concorrencia local.')
    }
    else {
        $classification = 'INCONCLUSIVO'
        $confidence = 'Media'
        [void]$findings.Add('Janela sem prova forte de quebra de transporte ou de colapso local.')
        [void]$actions.Add('Repetir a coleta durante o sintoma real.')
    }

    if ($null -ne $DbStats.P95) { [void]$findings.Add("DB p95 = $($DbStats.P95) ms") }
    if ($null -ne $NetStats.P95) { [void]$findings.Add("NetDir p95 = $($NetStats.P95) ms") }
    if ($null -ne $remoteToLocalRatio) { [void]$findings.Add("Relacao remoto/local = $remoteToLocalRatio x") }

    return [PSCustomObject]@{
        Classification = $classification
        Confidence = $confidence
        Findings = @($findings)
        Actions = @($actions)
        RemoteToLocalRatio = $remoteToLocalRatio
        RemoteAvgMs = $remoteAvg
    }
}

function Invoke-TargetProbe {
    param(
        $Target,
        [string]$ModeName,
        [string]$LocalProbePath,
        [string]$ExpectedExePath,
        [int]$Samples,
        [int]$IntervalSeconds,
        [int]$PingCount,
        [int]$TcpTimeoutMs
    )

    Log ("Iniciando alvo {0} | DB={1} | NetDir={2}" -f $Target.Label, $Target.DbPath, $Target.NetDirPath)

    if (-not (Test-Path -LiteralPath $LocalProbePath)) {
        Log ("LocalProbePath nao encontrado: {0}" -f $LocalProbePath) 'WARN'
    }

    $dbHostInfo = Resolve-HostInfo -HostName $Target.DbHost
    $netHostInfo = Resolve-HostInfo -HostName $Target.NetDirHost
    $dbPing = Test-HostPing -HostName $Target.DbHost -Count $PingCount
    $netPing = Test-HostPing -HostName $Target.NetDirHost -Count $PingCount
    $dbTcp = Test-TcpPortLatency -HostName $Target.DbHost -Port 445 -TimeoutMs $TcpTimeoutMs
    $netTcp = Test-TcpPortLatency -HostName $Target.NetDirHost -Port 445 -TimeoutMs $TcpTimeoutMs

    $startedAt = Get-Date
    $samplesData = @()
    for ($i = 1; $i -le $Samples; $i++) {
        $dbProbe = Measure-PathProbe -Path $Target.DbPath
        $netProbe = Measure-PathProbe -Path $Target.NetDirPath
        $localProbe = Measure-PathProbe -Path $LocalProbePath
        $cpu = Get-CpuPercentSafe

        $samplesData += [PSCustomObject]@{
            SampleIndex = $i
            Timestamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            CpuPercent = $cpu
            DbAccessible = $dbProbe.Accessible
            DbProbeMs = $dbProbe.ProbeMs
            DbError = $dbProbe.Error
            NetDirAccessible = $netProbe.Accessible
            NetDirProbeMs = $netProbe.ProbeMs
            NetDirError = $netProbe.Error
            LocalAccessible = $localProbe.Accessible
            LocalProbeMs = $localProbe.ProbeMs
            LocalError = $localProbe.Error
        }

        $cpuText = Get-TextOrDefault -Value $cpu -Default 'N/D' -Suffix '%'
        Log ("{0} | Amostra {1}/{2} | CPU={3} | DB={4} ({5} ms) | NetDir={6} ({7} ms) | Local={8} ({9} ms)" -f $Target.Label, $i, $Samples, $cpuText, $dbProbe.Accessible, $dbProbe.ProbeMs, $netProbe.Accessible, $netProbe.ProbeMs, $localProbe.Accessible, $localProbe.ProbeMs)

        if ($i -lt $Samples) { Start-Sleep -Seconds $IntervalSeconds }
    }
    $endedAt = Get-Date

    $dbStats = Get-Stats -Values @($samplesData | ForEach-Object { $_.DbProbeMs })
    $netStats = Get-Stats -Values @($samplesData | ForEach-Object { $_.NetDirProbeMs })
    $localStats = Get-Stats -Values @($samplesData | ForEach-Object { $_.LocalProbeMs })
    $cpuStats = Get-Stats -Values @($samplesData | ForEach-Object { $_.CpuPercent })
    $dbUnavailable = @($samplesData | Where-Object { -not $_.DbAccessible }).Count
    $netUnavailable = @($samplesData | Where-Object { -not $_.NetDirAccessible }).Count

    $smbSnapshot = Get-SmbSnapshot -Hosts @($Target.DbHost, $Target.NetDirHost)
    $dialects = @($smbSnapshot | ForEach-Object { $_.Dialect } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
    $servers = @($smbSnapshot | ForEach-Object { $_.ServerName } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)

    $assessment = Build-TargetAssessment -DbPing $dbPing -NetPing $netPing -DbTcp $dbTcp -NetTcp $netTcp -DbStats $dbStats -NetStats $netStats -LocalStats $localStats -DbUnavailable $dbUnavailable -NetUnavailable $netUnavailable -AvgCpu $cpuStats.Avg -PeakCpu $cpuStats.Max

    $compositeAvgMs = $null
    $compositeP95Ms = $null
    $valuesAvg = @()
    $valuesP95 = @()
    if ($null -ne $dbStats.Avg) { $valuesAvg += [double]$dbStats.Avg }
    if ($null -ne $netStats.Avg) { $valuesAvg += [double]$netStats.Avg }
    if ($null -ne $dbStats.P95) { $valuesP95 += [double]$dbStats.P95 }
    if ($null -ne $netStats.P95) { $valuesP95 += [double]$netStats.P95 }
    if ($valuesAvg.Count -gt 0) { $compositeAvgMs = [math]::Round((($valuesAvg | Measure-Object -Average).Average), 2) }
    if ($valuesP95.Count -gt 0) { $compositeP95Ms = [math]::Round((($valuesP95 | Measure-Object -Maximum).Maximum), 2) }

    return [PSCustomObject]@{
        Metadata = [PSCustomObject]@{
            ToolName = $script:ToolName
            ToolVersion = $script:ToolVersion
            RunId = $script:RunId
            Mode = $ModeName
            TargetLabel = $Target.Label
            StartedAt = $startedAt.ToString('yyyy-MM-dd HH:mm:ss')
            EndedAt = $endedAt.ToString('yyyy-MM-dd HH:mm:ss')
            Samples = $Samples
            IntervalSeconds = $IntervalSeconds
            ComputerName = $env:COMPUTERNAME
            UserName = $env:USERNAME
        }
        Target = [PSCustomObject]@{
            Label = $Target.Label
            DbPath = $Target.DbPath
            NetDirPath = $Target.NetDirPath
            DbHost = $Target.DbHost
            NetDirHost = $Target.NetDirHost
            HostOsHint = $Target.HostOsHint
            SmbDialectHint = $Target.SmbDialectHint
            LegacyHint = $Target.LegacyHint
            Notes = $Target.Notes
            ExpectedExePath = $ExpectedExePath
            ExeAccessible = (Test-Path -LiteralPath $ExpectedExePath)
            DbResolvedIPs = @($dbHostInfo.ResolvedIPs)
            NetDirResolvedIPs = @($netHostInfo.ResolvedIPs)
            ObservedSmbDialects = @($dialects)
            ObservedSmbServers = @($servers)
        }
        Network = [PSCustomObject]@{
            DbPing = $dbPing
            NetPing = $netPing
            DbTcp = $dbTcp
            NetTcp = $netTcp
            NetworkSnapshot = @(Get-NetworkSnapshot)
            SmbSnapshot = @($smbSnapshot)
        }
        Stats = [PSCustomObject]@{
            Db = $dbStats
            NetDir = $netStats
            Local = $localStats
            Cpu = $cpuStats
            DbUnavailableSamples = $dbUnavailable
            NetUnavailableSamples = $netUnavailable
            CompositeAvgMs = $compositeAvgMs
            CompositeP95Ms = $compositeP95Ms
        }
        Assessment = $assessment
        Timeline = @($samplesData)
    }
}

function Export-TargetTimelineCsv {
    param($TargetReport, [string]$Path)

    $rows = @(
        $TargetReport.Timeline |
        Select-Object -Property @('SampleIndex','Timestamp','CpuPercent','DbAccessible','DbProbeMs','NetDirAccessible','NetDirProbeMs','LocalAccessible','LocalProbeMs')
    )

    if ($rows.Count -eq 0) { return }
    $rows | Export-Csv -Path $Path -NoTypeInformation -Encoding UTF8
}


function Get-CompareBackendRankingGate {
    param($TargetReport)

    $dbDown = 0
    $netDown = 0
    $composite = $null

    try { if ($null -ne $TargetReport.Stats.DbUnavailableSamples) { $dbDown = [int]$TargetReport.Stats.DbUnavailableSamples } } catch {}
    try { if ($null -ne $TargetReport.Stats.NetUnavailableSamples) { $netDown = [int]$TargetReport.Stats.NetUnavailableSamples } } catch {}
    try { if ($null -ne $TargetReport.Stats.CompositeAvgMs) { $composite = [double]$TargetReport.Stats.CompositeAvgMs } } catch {}

    $eligible = ($dbDown -eq 0 -and $netDown -eq 0 -and $null -ne $composite)
    $gateStatus = if ($eligible) { 'ELEGIVEL' } else { 'DESCLASSIFICADO' }

    $reasons = New-Object System.Collections.Generic.List[string]
    if ($eligible) {
        [void]$reasons.Add('Sem indisponibilidade de DB/NetDir na janela comparativa.')
    }
    else {
        if ($dbDown -gt 0) { [void]$reasons.Add("DB indisponivel em $dbDown amostra(s)") }
        if ($netDown -gt 0) { [void]$reasons.Add("NetDir indisponivel em $netDown amostra(s)") }
        if ($null -eq $composite) { [void]$reasons.Add('Composto sem valor calculavel') }
    }

    return [PSCustomObject]@{
        Eligible = $eligible
        GateStatus = $gateStatus
        GateReason = (@($reasons) -join '; ')
        DbDown = $dbDown
        NetDown = $netDown
        EffectiveCompositeAvgMs = $(if ($eligible) { $composite } else { $null })
    }
}

function Build-CompareBackendSummary {
    param([object[]]$TargetReports, [string]$WorkloadLabel)

    $annotatedReports = @()
    foreach ($report in @($TargetReports)) {
        try {
            $report | Add-Member -NotePropertyName RankingGate -NotePropertyValue (Get-CompareBackendRankingGate -TargetReport $report) -Force
        }
        catch {}
        $annotatedReports += $report
    }

    $reports = @($annotatedReports | Sort-Object -Property `
        @{ Expression = { if ($_.RankingGate -and $_.RankingGate.Eligible) { 0 } else { 1 } }; Ascending = $true }, `
        @{ Expression = { if ($_.RankingGate -and $null -ne $_.RankingGate.EffectiveCompositeAvgMs) { [double]$_.RankingGate.EffectiveCompositeAvgMs } else { 999999 } }; Ascending = $true })

    $eligibleReports = @($reports | Where-Object { $_.RankingGate -and $_.RankingGate.Eligible })
    $best = $null
    if ($eligibleReports.Count -gt 0) { $best = $eligibleReports[0] }

    $legacyReports = @($reports | Where-Object { $_.Target.LegacyHint })
    $modernReports = @($reports | Where-Object { -not $_.Target.LegacyHint })

    $legacyEligible = @($legacyReports | Where-Object { $_.RankingGate -and $_.RankingGate.Eligible })
    $modernEligible = @($modernReports | Where-Object { $_.RankingGate -and $_.RankingGate.Eligible })

    $legacyBest = $null
    $modernBest = $null
    if ($legacyEligible.Count -gt 0) { $legacyBest = $legacyEligible[0] }
    if ($modernEligible.Count -gt 0) { $modernBest = $modernEligible[0] }

    $correlation = 'N/D'
    $summary = 'Nao ha dados suficientes para inferir correlacao com compatibilidade legada.'
    $recommendation = 'Executar nova rodada com pelo menos um alvo legado e um alvo moderno habilitados.'
    $deltaMs = $null
    $ratio = $null
    $comparisonValidity = 'OK'
    $comparisonValidityReason = 'Todos os alvos comparados elegiveis para ranking.'

    if ($legacyReports.Count -gt 0 -and $legacyEligible.Count -eq 0) {
        $comparisonValidity = 'INVALIDA'
        $comparisonValidityReason = 'O alvo legado ficou desclassificado por indisponibilidade de DB/NetDir na janela.'
    }
    elseif ($modernReports.Count -gt 0 -and $modernEligible.Count -eq 0) {
        $comparisonValidity = 'INVALIDA'
        $comparisonValidityReason = 'O alvo moderno ficou desclassificado por indisponibilidade de DB/NetDir na janela.'
    }
    elseif ($legacyEligible.Count -eq 0 -and $modernEligible.Count -eq 0 -and ($legacyReports.Count -gt 0 -or $modernReports.Count -gt 0)) {
        $comparisonValidity = 'INVALIDA'
        $comparisonValidityReason = 'Todos os alvos ficaram desclassificados por indisponibilidade ou ausencia de composto valido.'
    }

    if ($comparisonValidity -eq 'INVALIDA') {
        $correlation = 'INVALIDA'
        $summary = 'A janela comparativa ficou invalida para conclusao final: pelo menos um dos alvos apresentou indisponibilidade de DB/NetDir e nao pode ser ranqueado de forma justa.'
        $recommendation = 'Repetir a coleta somente quando todos os alvos estiverem elegiveis. Nao considerar vencedor quando houver alvo desclassificado por indisponibilidade.'
    }
    elseif ($legacyBest -and $modernBest -and $null -ne $legacyBest.Stats.CompositeAvgMs -and $null -ne $modernBest.Stats.CompositeAvgMs) {
        $deltaMs = [math]::Round(([double]$modernBest.Stats.CompositeAvgMs - [double]$legacyBest.Stats.CompositeAvgMs), 2)
        if ([double]$legacyBest.Stats.CompositeAvgMs -gt 0) {
            $ratio = [math]::Round(([double]$modernBest.Stats.CompositeAvgMs / [double]$legacyBest.Stats.CompositeAvgMs), 2)
        }

        if ($deltaMs -ge 15 -and $ratio -ge 1.5) {
            $correlation = 'FORTE'
            $summary = "O backend legado apresentou latencia composta significativamente menor que o backend moderno para $WorkloadLabel."
            $recommendation = 'A hipotese de compatibilidade DBE/BDE com stack legada ganha forca. Validar se o backend moderno introduz custo de protocolo, locking ou antivirus.'
        }
        elseif ($deltaMs -ge 8 -and $ratio -ge 1.2) {
            $correlation = 'MEDIA'
            $summary = "O backend legado ficou melhor que o moderno, mas a diferenca ainda pede repeticao controlada para fechar causalidade."
            $recommendation = 'Repetir a coleta no mesmo posto/mesma carga e confrontar com uso real do ECG.'
        }
        elseif ([math]::Abs($deltaMs) -lt 8) {
            $correlation = 'FRACA'
            $summary = 'Os backends performaram de forma parecida; a hipotese de compatibilidade legada perde forca nesta janela.'
            $recommendation = 'Investigar rede, storage, antivirus e lock contention antes de atribuir o problema ao protocolo.'
        }
        elseif ($deltaMs -le -8) {
            $correlation = 'CONTRARIA'
            $summary = 'O backend moderno performou melhor que o legado; a hipotese de compatibilidade SMB1 como fator positivo nao se sustentou nesta janela.'
            $recommendation = 'Revisar a tese de legado e priorizar investigacao no proprio host XP/SMB1 ou em gargalos externos.'
        }
    }

    $rows = @()
    foreach ($r in @($reports)) {
        $rows += [PSCustomObject]@{
            Alvo = $r.Target.Label
            Legacy = $(if ($r.Target.LegacyHint) { 'Sim' } else { 'Nao' })
            HostOsHint = $r.Target.HostOsHint
            SmbHint = $r.Target.SmbDialectHint
            SmbObservado = @($r.Target.ObservedSmbDialects) -join ', '
            Classificacao = $r.Assessment.Classification
            Confianca = $r.Assessment.Confidence
            ElegivelRanking = $(if ($r.RankingGate.Eligible) { 'Sim' } else { 'Nao' })
            GateStatus = $r.RankingGate.GateStatus
            GateReason = $r.RankingGate.GateReason
            MediaDBms = $r.Stats.Db.Avg
            P95DBms = $r.Stats.Db.P95
            MediaNetDirMs = $r.Stats.NetDir.Avg
            P95NetDirMs = $r.Stats.NetDir.P95
            CompositeAvgMs = $r.Stats.CompositeAvgMs
            CompositeP95Ms = $r.Stats.CompositeP95Ms
            CompositeRankMs = $r.RankingGate.EffectiveCompositeAvgMs
            RazaoRemotoLocal = $r.Assessment.RemoteToLocalRatio
            DbDown = $r.Stats.DbUnavailableSamples
            NetDown = $r.Stats.NetUnavailableSamples
        }
    }

    return [PSCustomObject]@{
        BestTargetLabel = $(if ($best) { $best.Target.Label } else { 'N/D' })
        LegacyCorrelation = $correlation
        Summary = $summary
        Recommendation = $recommendation
        DeltaMsModernVsLegacy = $deltaMs
        RatioModernVsLegacy = $ratio
        ComparisonValidity = $comparisonValidity
        ComparisonValidityReason = $comparisonValidityReason
        Rows = @($rows)
    }
}

function New-CompareBackendHtml {
    param($Summary, [object[]]$TargetReports, [string]$WorkloadLabel)

    $badgeClass = 'inconclusivo'
    switch ($Summary.LegacyCorrelation) {
        'FORTE' { $badgeClass = 'critico' }
        'MEDIA' { $badgeClass = 'lento' }
        'FRACA' { $badgeClass = 'normal' }
        'CONTRARIA' { $badgeClass = 'normal' }
        'INVALIDA' { $badgeClass = 'inconclusivo' }
    }

    $summaryTable = ConvertTo-HtmlTable -Rows @($Summary.Rows)

    $targetSections = New-Object System.Text.StringBuilder
    foreach ($r in @($TargetReports)) {
        $smbObserved = @($r.Target.ObservedSmbDialects) -join ', '
        if ([string]::IsNullOrWhiteSpace($smbObserved)) { $smbObserved = 'N/D' }

        $targetRows = @(
            [PSCustomObject]@{ Campo = 'DB'; Valor = $r.Target.DbPath },
            [PSCustomObject]@{ Campo = 'NetDir'; Valor = $r.Target.NetDirPath },
            [PSCustomObject]@{ Campo = 'Host OS hint'; Valor = $r.Target.HostOsHint },
            [PSCustomObject]@{ Campo = 'SMB hint'; Valor = $r.Target.SmbDialectHint },
            [PSCustomObject]@{ Campo = 'SMB observado'; Valor = $smbObserved },
            [PSCustomObject]@{ Campo = 'Classificacao'; Valor = $r.Assessment.Classification },
            [PSCustomObject]@{ Campo = 'Elegivel para ranking'; Valor = $(if ($r.RankingGate.Eligible) { 'Sim' } else { 'Nao' }) },
            [PSCustomObject]@{ Campo = 'Gate'; Valor = $r.RankingGate.GateReason },
            [PSCustomObject]@{ Campo = 'Resumo'; Valor = (@($r.Assessment.Findings) -join ' | ') }
        )
        $statsRows = @(
            [PSCustomObject]@{ Metrica = 'DB media/p95/pico'; Valor = ([string]$r.Stats.Db.Avg + ' / ' + [string]$r.Stats.Db.P95 + ' / ' + [string]$r.Stats.Db.Max) },
            [PSCustomObject]@{ Metrica = 'NetDir media/p95/pico'; Valor = ([string]$r.Stats.NetDir.Avg + ' / ' + [string]$r.Stats.NetDir.P95 + ' / ' + [string]$r.Stats.NetDir.Max) },
            [PSCustomObject]@{ Metrica = 'Local media/p95/pico'; Valor = ([string]$r.Stats.Local.Avg + ' / ' + [string]$r.Stats.Local.P95 + ' / ' + [string]$r.Stats.Local.Max) },
            [PSCustomObject]@{ Metrica = 'Composto media/p95'; Valor = ([string]$r.Stats.CompositeAvgMs + ' / ' + [string]$r.Stats.CompositeP95Ms) },
            [PSCustomObject]@{ Metrica = 'Composto elegivel'; Valor = $(if ($null -ne $r.RankingGate.EffectiveCompositeAvgMs) { [string]$r.RankingGate.EffectiveCompositeAvgMs } else { 'N/D' }) },
            [PSCustomObject]@{ Metrica = 'Razao remoto/local'; Valor = $r.Assessment.RemoteToLocalRatio },
            [PSCustomObject]@{ Metrica = 'DB/NetDir indisponivel'; Valor = ([string]$r.Stats.DbUnavailableSamples + ' / ' + [string]$r.Stats.NetUnavailableSamples) }
        )
        [void]$targetSections.Append('<div class="card">')
        [void]$targetSections.Append('<h2>' + [System.Net.WebUtility]::HtmlEncode($r.Target.Label) + '</h2>')
        [void]$targetSections.Append('<div class="grid">')
        [void]$targetSections.Append('<div>' + (ConvertTo-HtmlTable -Rows $targetRows) + '</div>')
        [void]$targetSections.Append('<div>' + (ConvertTo-HtmlTable -Rows $statsRows) + '</div>')
        [void]$targetSections.Append('</div>')
        [void]$targetSections.Append('</div>')
    }

    @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>ECG CompareBackend Report - v6.3.2</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f6fb;color:#1f2937}
.wrapper{max-width:1260px;margin:0 auto;padding:24px}
.card{background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(15,23,42,.08);margin-bottom:16px;padding:20px}
.grid{display:grid;gap:12px;grid-template-columns:repeat(2,1fr)}
.kv{border:1px solid #e5e7eb;border-radius:12px;padding:10px 12px;background:#fafbfd}.kv strong{display:block;font-size:12px;color:#6b7280}
.badge{display:inline-block;padding:8px 14px;border-radius:999px;font-weight:700;font-size:13px}
.badge.normal{background:#dcfce7;color:#166534}.badge.lento{background:#fef3c7;color:#92400e}.badge.critico{background:#fee2e2;color:#991b1b}.badge.inconclusivo{background:#e5e7eb;color:#374151}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #e5e7eb;padding:8px 10px;text-align:left;vertical-align:top}th{background:#f9fafb}
.muted{color:#6b7280;font-size:12px}
@media (max-width:960px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrapper">
<div class="card">
<div class="badge $badgeClass">$($Summary.LegacyCorrelation)</div>
<h1>ECG Diagnostics Core - CompareBackend</h1>
<p class="muted">Build patched: $([System.Net.WebUtility]::HtmlEncode($script:ToolVersion))</p>
<p>$([System.Net.WebUtility]::HtmlEncode($Summary.Summary))</p>
<div class="grid">
<div class="kv"><strong>Workload</strong>$([System.Net.WebUtility]::HtmlEncode($WorkloadLabel))</div>
<div class="kv"><strong>Melhor alvo</strong>$([System.Net.WebUtility]::HtmlEncode([string]$Summary.BestTargetLabel))</div>
<div class="kv"><strong>Delta moderno vs legado</strong>$([System.Net.WebUtility]::HtmlEncode([string]$Summary.DeltaMsModernVsLegacy)) ms</div>
<div class="kv"><strong>Razao moderno / legado</strong>$([System.Net.WebUtility]::HtmlEncode([string]$Summary.RatioModernVsLegacy)) x</div>
<div class="kv"><strong>Validade da comparacao</strong>$([System.Net.WebUtility]::HtmlEncode([string]$Summary.ComparisonValidity))</div>
<div class="kv"><strong>Motivo da validade</strong>$([System.Net.WebUtility]::HtmlEncode([string]$Summary.ComparisonValidityReason))</div>
<div class="kv"><strong>Recomendacao</strong>$([System.Net.WebUtility]::HtmlEncode($Summary.Recommendation))</div>
<div class="kv"><strong>RunId</strong>$([System.Net.WebUtility]::HtmlEncode($script:RunId))</div>
</div>
</div>
<div class="card"><h2>Resumo comparativo</h2>$summaryTable</div>
$($targetSections.ToString())
</div>
</body>
</html>
"@
}

function New-SingleHtml {
    param($TargetReport, [string]$WorkloadLabel)

    $badgeClass = 'inconclusivo'
    switch ($TargetReport.Assessment.Classification) {
        'REDE/CAMINHO' { $badgeClass = 'critico' }
        'SERVIDOR/SHARE' { $badgeClass = 'lento' }
        'ESTACAO LOCAL' { $badgeClass = 'lento' }
        'INCONCLUSIVO' { $badgeClass = 'inconclusivo' }
    }

    $overviewRows = @(
        [PSCustomObject]@{ Campo = 'Alvo'; Valor = $TargetReport.Target.Label },
        [PSCustomObject]@{ Campo = 'DB'; Valor = $TargetReport.Target.DbPath },
        [PSCustomObject]@{ Campo = 'NetDir'; Valor = $TargetReport.Target.NetDirPath },
        [PSCustomObject]@{ Campo = 'Host OS hint'; Valor = $TargetReport.Target.HostOsHint },
        [PSCustomObject]@{ Campo = 'SMB hint'; Valor = $TargetReport.Target.SmbDialectHint },
        [PSCustomObject]@{ Campo = 'SMB observado'; Valor = (@($TargetReport.Target.ObservedSmbDialects) -join ', ') }
    )
    $statsRows = @(
        [PSCustomObject]@{ Metrica = 'DB media/p95/pico'; Valor = ("{0} / {1} / {2}" -f $TargetReport.Stats.Db.Avg, $TargetReport.Stats.Db.P95, $TargetReport.Stats.Db.Max) },
        [PSCustomObject]@{ Metrica = 'NetDir media/p95/pico'; Valor = ("{0} / {1} / {2}" -f $TargetReport.Stats.NetDir.Avg, $TargetReport.Stats.NetDir.P95, $TargetReport.Stats.NetDir.Max) },
        [PSCustomObject]@{ Metrica = 'Local media/p95/pico'; Valor = ("{0} / {1} / {2}" -f $TargetReport.Stats.Local.Avg, $TargetReport.Stats.Local.P95, $TargetReport.Stats.Local.Max) },
        [PSCustomObject]@{ Metrica = 'Composto media/p95'; Valor = ("{0} / {1}" -f $TargetReport.Stats.CompositeAvgMs, $TargetReport.Stats.CompositeP95Ms) },
        [PSCustomObject]@{ Metrica = 'CPU media/pico'; Valor = ("{0} / {1}" -f $TargetReport.Stats.Cpu.Avg, $TargetReport.Stats.Cpu.Max) },
        [PSCustomObject]@{ Metrica = 'DB/NetDir indisponivel'; Valor = ("{0} / {1}" -f $TargetReport.Stats.DbUnavailableSamples, $TargetReport.Stats.NetUnavailableSamples) }
    )
    $smbRows = @($TargetReport.Network.SmbSnapshot)
    $timelineRows = @(
        $TargetReport.Timeline |
        Select-Object -First 20 -Property @('SampleIndex','Timestamp','CpuPercent','DbAccessible','DbProbeMs','NetDirAccessible','NetDirProbeMs','LocalProbeMs')
    )

    @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>ECG Single Report - v6.3.2</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f6fb;color:#1f2937}
.wrapper{max-width:1260px;margin:0 auto;padding:24px}
.card{background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(15,23,42,.08);margin-bottom:16px;padding:20px}
.grid{display:grid;gap:12px;grid-template-columns:repeat(2,1fr)}
.badge{display:inline-block;padding:8px 14px;border-radius:999px;font-weight:700;font-size:13px}
.badge.normal{background:#dcfce7;color:#166534}.badge.lento{background:#fef3c7;color:#92400e}.badge.critico{background:#fee2e2;color:#991b1b}.badge.inconclusivo{background:#e5e7eb;color:#374151}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #e5e7eb;padding:8px 10px;text-align:left;vertical-align:top}th{background:#f9fafb}
.muted{color:#6b7280;font-size:12px}
@media (max-width:960px){.grid{grid-template-columns:1fr}}
</style>
</head>
<body>
<div class="wrapper">
<div class="card">
<div class="badge $badgeClass">$($TargetReport.Assessment.Classification)</div>
<h1>ECG Diagnostics Core - Single</h1>
<p class="muted">Build patched: $([System.Net.WebUtility]::HtmlEncode($script:ToolVersion))</p>
<p>$([System.Net.WebUtility]::HtmlEncode((@($TargetReport.Assessment.Findings) -join ' | ')))</p>
<p class="muted">Workload: $([System.Net.WebUtility]::HtmlEncode($WorkloadLabel)) | RunId: $([System.Net.WebUtility]::HtmlEncode($TargetReport.Metadata.RunId))</p>
</div>
<div class="card"><h2>Alvo</h2><div class="grid"><div>$(ConvertTo-HtmlTable -Rows $overviewRows)</div><div>$(ConvertTo-HtmlTable -Rows $statsRows)</div></div></div>
<div class="card"><h2>Rede / SMB</h2>$(ConvertTo-HtmlTable -Rows @([PSCustomObject]@{ Alvo='DB'; Host=$TargetReport.Target.DbHost; PingMedioMs=$TargetReport.Network.DbPing.AvgMs; PingPerda=$TargetReport.Network.DbPing.Lost; Tcp445=$(if($TargetReport.Network.DbTcp.Reachable){'OK'}else{'FALHA'}); TcpMs=$TargetReport.Network.DbTcp.ConnectMs },[PSCustomObject]@{ Alvo='NetDir'; Host=$TargetReport.Target.NetDirHost; PingMedioMs=$TargetReport.Network.NetPing.AvgMs; PingPerda=$TargetReport.Network.NetPing.Lost; Tcp445=$(if($TargetReport.Network.NetTcp.Reachable){'OK'}else{'FALHA'}); TcpMs=$TargetReport.Network.NetTcp.ConnectMs }))</div>
<div class="card"><h2>Conexoes SMB observadas</h2>$(ConvertTo-HtmlTable -Rows $smbRows)</div>
<div class="card"><h2>Timeline (primeiras 20 amostras)</h2>$(ConvertTo-HtmlTable -Rows $timelineRows)</div>
</div>
</body>
</html>
"@
}

function Build-CompareBackendJson {
    param($Summary, [object[]]$TargetReports, [string]$WorkloadLabel)
    return [PSCustomObject]@{
        Metadata = [PSCustomObject]@{
            ToolName = $script:ToolName
            ToolVersion = $script:ToolVersion
            RunId = $script:RunId
            Mode = 'CompareBackend'
            GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            ComputerName = $env:COMPUTERNAME
            UserName = $env:USERNAME
            WorkloadLabel = $WorkloadLabel
            CoreScriptPath = $script:ScriptPath
            ProfilePath = $ProfilePath
        }
        Comparison = $Summary
        Targets = @($TargetReports)
        Logs = @($script:LogLines)
    }
}

function Build-SingleJson {
    param($TargetReport, [string]$WorkloadLabel)
    return [PSCustomObject]@{
        Metadata = [PSCustomObject]@{
            ToolName = $script:ToolName
            ToolVersion = $script:ToolVersion
            RunId = $script:RunId
            Mode = 'Single'
            GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
            ComputerName = $env:COMPUTERNAME
            UserName = $env:USERNAME
            WorkloadLabel = $WorkloadLabel
            CoreScriptPath = $script:ScriptPath
            ProfilePath = $ProfilePath
        }
        TargetReport = $TargetReport
        Logs = @($script:LogLines)
    }
}

function Build-CompareJsonReport {
    param($LeftData, $RightData, [string]$LeftPath, [string]$RightPath)

    $rows = @()
    $leftTargets = @()
    $rightTargets = @()

    if ($LeftData.Targets) { $leftTargets = @($LeftData.Targets) }
    elseif ($LeftData.TargetReport) { $leftTargets = @($LeftData.TargetReport) }

    if ($RightData.Targets) { $rightTargets = @($RightData.Targets) }
    elseif ($RightData.TargetReport) { $rightTargets = @($RightData.TargetReport) }

    foreach ($lt in @($leftTargets)) {
        $rightMatch = $null
        foreach ($rt in @($rightTargets)) {
            if ([string]$rt.Target.Label -eq [string]$lt.Target.Label) { $rightMatch = $rt; break }
        }

        $rows += [PSCustomObject]@{
            Alvo = $lt.Target.Label
            LeftCompositeAvgMs = $lt.Stats.CompositeAvgMs
            RightCompositeAvgMs = $(if ($rightMatch) { $rightMatch.Stats.CompositeAvgMs } else { $null })
            LeftClassificacao = $lt.Assessment.Classification
            RightClassificacao = $(if ($rightMatch) { $rightMatch.Assessment.Classification } else { 'N/D' })
            Resultado = $(if ($rightMatch) {
                if ($lt.Stats.CompositeAvgMs -eq $rightMatch.Stats.CompositeAvgMs) { 'IGUAL' } else { 'DIFERENTE' }
            } else { 'AUSENTE' })
        }
    }

    $html = @"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<title>ECG CompareJson Report - v6.3.2</title>
<style>
body{font-family:Segoe UI,Arial,sans-serif;margin:0;background:#f3f6fb;color:#1f2937}
.wrapper{max-width:1200px;margin:0 auto;padding:24px}
.card{background:#fff;border-radius:14px;box-shadow:0 2px 8px rgba(15,23,42,.08);margin-bottom:16px;padding:20px}
table{width:100%;border-collapse:collapse;font-size:13px}th,td{border:1px solid #e5e7eb;padding:8px 10px;text-align:left;vertical-align:top}th{background:#f9fafb}
.muted{color:#6b7280;font-size:12px}
</style>
</head>
<body>
<div class="wrapper">
<div class="card">
<h1>ECG Diagnostics Core - CompareJson</h1>
<p class="muted">Build patched: $([System.Net.WebUtility]::HtmlEncode($script:ToolVersion))</p>
<p class="muted"><strong>Esquerda:</strong> $([System.Net.WebUtility]::HtmlEncode($LeftPath))<br><strong>Direita:</strong> $([System.Net.WebUtility]::HtmlEncode($RightPath))</p>
</div>
<div class="card">
<h2>Comparacao</h2>
$(ConvertTo-HtmlTable -Rows $rows)
</div>
</div>
</body>
</html>
"@

    return [PSCustomObject]@{
        Rows = @($rows)
        Html = $html
    }
}

function Write-TargetSummaryText {
    param($TargetReport, [string]$Path)

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("$($script:ToolName) $($script:ToolVersion)")
    [void]$lines.Add("RunId: $($TargetReport.Metadata.RunId)")
    [void]$lines.Add("Alvo: $($TargetReport.Target.Label)")
    [void]$lines.Add("Classificacao: $($TargetReport.Assessment.Classification)")
    [void]$lines.Add("Confianca: $($TargetReport.Assessment.Confidence)")
    [void]$lines.Add('')
    [void]$lines.Add('Achados:')
    foreach ($line in @($TargetReport.Assessment.Findings)) { [void]$lines.Add('- ' + [string]$line) }
    [void]$lines.Add('')
    [void]$lines.Add('Proximas acoes:')
    foreach ($line in @($TargetReport.Assessment.Actions)) { [void]$lines.Add('- ' + [string]$line) }
    [void]$lines.Add('')
    [void]$lines.Add("DB media/p95/pico: $($TargetReport.Stats.Db.Avg) / $($TargetReport.Stats.Db.P95) / $($TargetReport.Stats.Db.Max) ms")
    [void]$lines.Add("NetDir media/p95/pico: $($TargetReport.Stats.NetDir.Avg) / $($TargetReport.Stats.NetDir.P95) / $($TargetReport.Stats.NetDir.Max) ms")
    [void]$lines.Add("Local media/p95/pico: $($TargetReport.Stats.Local.Avg) / $($TargetReport.Stats.Local.P95) / $($TargetReport.Stats.Local.Max) ms")
    [void]$lines.Add("SMB observado: " + ((@($TargetReport.Target.ObservedSmbDialects) -join ', ')))
    Write-Utf8File -Path $Path -Text ((@($lines) -join [Environment]::NewLine))
}

function Write-CompareSummaryText {
    param($Summary, [string]$Path)

    $lines = New-Object System.Collections.Generic.List[string]
    [void]$lines.Add("$($script:ToolName) $($script:ToolVersion)")
    [void]$lines.Add("RunId: $($script:RunId)")
    [void]$lines.Add("Correlacao legada: $($Summary.LegacyCorrelation)")
    [void]$lines.Add("Melhor alvo: $($Summary.BestTargetLabel)")
    [void]$lines.Add("Validade da comparacao: $($Summary.ComparisonValidity)")
    [void]$lines.Add("Motivo da validade: $($Summary.ComparisonValidityReason)")
    [void]$lines.Add("Resumo: $($Summary.Summary)")
    [void]$lines.Add("Recomendacao: $($Summary.Recommendation)")
    [void]$lines.Add('')
    [void]$lines.Add('Ranking:')
    foreach ($row in @($Summary.Rows)) {
        $line = '- ' + [string]$row.Alvo +
            ' | Legacy=' + [string]$row.Legacy +
            ' | Elegivel=' + [string]$row.ElegivelRanking +
            ' | Gate=' + [string]$row.GateStatus +
            ' | CompositeAvgMs=' + [string]$row.CompositeAvgMs +
            ' | CompositeRankMs=' + [string]$row.CompositeRankMs +
            ' | Classificacao=' + [string]$row.Classificacao +
            ' | SMB obs=' + [string]$row.SmbObservado
        [void]$lines.Add($line)
    }
    Write-Utf8File -Path $Path -Text ((@($lines) -join [Environment]::NewLine))
}

function Invoke-SingleMode {
    param($Ini)

    $targets = @(Get-EnabledTargets -Ini $Ini)
    if ($targets.Count -eq 0) { throw 'Nenhum target habilitado e valido no INI.' }

    $target = $null
    if (-not [string]::IsNullOrWhiteSpace($TargetLabel)) {
        foreach ($t in $targets) {
            if ([string]$t.Label -eq [string]$TargetLabel) { $target = $t; break }
        }
        if ($null -eq $target) { throw "TargetLabel nao encontrado no INI: $TargetLabel" }
    }
    else {
        $target = $targets[0]
    }

    $runRoot = Join-Path (Resolve-CompareOutDir -Ini $Ini) $script:RunId
    $script:CurrentRunRoot = $runRoot
    Ensure-Directory -Path $runRoot

    $localProbePath = Normalize-PathString (Get-IniString -Ini $Ini -Section 'General' -Key 'LocalProbePath' -Default 'C:\Windows')
    $expectedExePath = Normalize-PathString (Get-IniString -Ini $Ini -Section 'General' -Key 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe')
    $workloadLabel = Get-IniString -Ini $Ini -Section 'General' -Key 'WorkloadLabel' -Default 'ECGv6 DBE/BDE'
    $samples = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'Samples' -Default 20))
    $interval = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'IntervalSeconds' -Default 3))
    $pingCount = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'PingCount' -Default 4))
    $tcpTimeout = [Math]::Max(100, (Get-IniInt -Ini $Ini -Section 'General' -Key 'TcpTimeoutMs' -Default 1500))

    Log ("Single target selecionado: {0} ({1})" -f $target.Label, $target.DbPath) 'STEP'

    $targetReport = Invoke-TargetProbe -Target $target -ModeName 'Single' -LocalProbePath $localProbePath -ExpectedExePath $expectedExePath -Samples $samples -IntervalSeconds $interval -PingCount $pingCount -TcpTimeoutMs $tcpTimeout

    $html = New-SingleHtml -TargetReport $targetReport -WorkloadLabel $workloadLabel
    $json = Build-SingleJson -TargetReport $targetReport -WorkloadLabel $workloadLabel

    $htmlPath = Join-Path $runRoot 'Single_Report.html'
    $jsonPath = Join-Path $runRoot 'Single_Report.json'
    $txtPath = Join-Path $runRoot 'Single_Summary.txt'
    $csvPath = Join-Path $runRoot ("Single_{0}.csv" -f $target.Label)

    Write-Utf8File -Path $htmlPath -Text $html
    Write-Utf8File -Path $jsonPath -Text ($json | ConvertTo-Json -Depth 10)
    Write-TargetSummaryText -TargetReport $targetReport -Path $txtPath
    Export-TargetTimelineCsv -TargetReport $targetReport -Path $csvPath

    Log "HTML salvo em $htmlPath" 'STEP'
    Log "JSON salvo em $jsonPath" 'STEP'
    Log "TXT salvo em $txtPath" 'STEP'
    Log "CSV salvo em $csvPath" 'STEP'

    if ($OpenReport) {
        try { Start-Process $htmlPath | Out-Null } catch {}
    }
}

function Invoke-CompareBackendMode {
    param($Ini)

    $targets = @(Get-EnabledTargets -Ini $Ini)
    if ($targets.Count -lt 2) { throw 'CompareBackend requer pelo menos 2 targets habilitados e validos no INI.' }

    $runRoot = Join-Path (Resolve-CompareOutDir -Ini $Ini) $script:RunId
    $script:CurrentRunRoot = $runRoot
    Ensure-Directory -Path $runRoot

    $localProbePath = Normalize-PathString (Get-IniString -Ini $Ini -Section 'General' -Key 'LocalProbePath' -Default 'C:\Windows')
    $expectedExePath = Normalize-PathString (Get-IniString -Ini $Ini -Section 'General' -Key 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe')
    $workloadLabel = Get-IniString -Ini $Ini -Section 'General' -Key 'WorkloadLabel' -Default 'ECGv6 DBE/BDE'
    $samples = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'Samples' -Default 20))
    $interval = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'IntervalSeconds' -Default 3))
    $pingCount = [Math]::Max(1, (Get-IniInt -Ini $Ini -Section 'General' -Key 'PingCount' -Default 4))
    $tcpTimeout = [Math]::Max(100, (Get-IniInt -Ini $Ini -Section 'General' -Key 'TcpTimeoutMs' -Default 1500))

    $reports = @()
    foreach ($target in @($targets)) {
        $report = Invoke-TargetProbe -Target $target -ModeName 'CompareBackend' -LocalProbePath $localProbePath -ExpectedExePath $expectedExePath -Samples $samples -IntervalSeconds $interval -PingCount $pingCount -TcpTimeoutMs $tcpTimeout
        $reports += $report

        $csvPath = Join-Path $runRoot ("CompareBackend_{0}.csv" -f $target.Label)
        Export-TargetTimelineCsv -TargetReport $report -Path $csvPath
        Log "CSV salvo em $csvPath" 'STEP'
    }

    $summary = Build-CompareBackendSummary -TargetReports $reports -WorkloadLabel $workloadLabel
    $html = New-CompareBackendHtml -Summary $summary -TargetReports $reports -WorkloadLabel $workloadLabel
    $json = Build-CompareBackendJson -Summary $summary -TargetReports $reports -WorkloadLabel $workloadLabel

    $htmlPath = Join-Path $runRoot 'CompareBackend_Report.html'
    $jsonPath = Join-Path $runRoot 'CompareBackend_Report.json'
    $txtPath = Join-Path $runRoot 'CompareBackend_Summary.txt'

    Write-Utf8File -Path $htmlPath -Text $html
    Write-Utf8File -Path $jsonPath -Text ($json | ConvertTo-Json -Depth 10)
    Write-CompareSummaryText -Summary $summary -Path $txtPath

    Log "HTML salvo em $htmlPath" 'STEP'
    Log "JSON salvo em $jsonPath" 'STEP'
    Log "TXT salvo em $txtPath" 'STEP'

    if ($OpenReport) {
        try { Start-Process $htmlPath | Out-Null } catch {}
    }
}

function Invoke-CompareJsonMode {
    if ([string]::IsNullOrWhiteSpace($CompareLeftReport) -or [string]::IsNullOrWhiteSpace($CompareRightReport)) {
        throw 'CompareJson requer -CompareLeftReport e -CompareRightReport.'
    }
    if (-not (Test-Path -LiteralPath $CompareLeftReport)) { throw "Arquivo esquerdo nao encontrado: $CompareLeftReport" }
    if (-not (Test-Path -LiteralPath $CompareRightReport)) { throw "Arquivo direito nao encontrado: $CompareRightReport" }

    $leftData = Get-Content -LiteralPath $CompareLeftReport -Raw | ConvertFrom-Json
    $rightData = Get-Content -LiteralPath $CompareRightReport -Raw | ConvertFrom-Json
    $report = Build-CompareJsonReport -LeftData $leftData -RightData $rightData -LeftPath $CompareLeftReport -RightPath $CompareRightReport

    $runRoot = Join-Path (Resolve-CompareOutDir -Ini (Read-IniFile -Path $ProfilePath)) $script:RunId
    $script:CurrentRunRoot = $runRoot
    Ensure-Directory -Path $runRoot

    $htmlPath = Join-Path $runRoot 'CompareJson_Report.html'
    $jsonPath = Join-Path $runRoot 'CompareJson_Report.json'
    Write-Utf8File -Path $htmlPath -Text $report.Html
    Write-Utf8File -Path $jsonPath -Text ($report.Rows | ConvertTo-Json -Depth 8)

    Log "HTML salvo em $htmlPath" 'STEP'
    Log "JSON salvo em $jsonPath" 'STEP'

    if ($OpenReport) {
        try { Start-Process $htmlPath | Out-Null } catch {}
    }
}



try {
    switch ($Mode) {
        'Fix'            { Invoke-FixAutoMode -ApplyFixes $true }
        'Auto'           { Invoke-FixAutoMode -ApplyFixes $false }
        'Detect'         { Invoke-FixAutoMode -ApplyFixes $false }
        'Compare'        { Invoke-CompareMode }
        'CompareJson'    { Invoke-CompareJsonMode }
        'Rollback'       { Invoke-RollbackMode }
        'Monitor'        { Invoke-MonitorMode }
        'CollectStatic'  { Invoke-CollectStaticMode }
        'Single'         {
            if ([string]::IsNullOrWhiteSpace($ProfilePath)) { throw 'ProfilePath nao resolvido.' }
            $ini = Read-IniFile -Path $ProfilePath
            Log "$($script:ToolName) $($script:ToolVersion)"
            Log "RunId: $($script:RunId)"
            Log "Mode: $Mode"
            Log "Profile: $ProfilePath"
            Invoke-SingleMode -Ini $ini
        }
        'CompareBackend' {
            if ([string]::IsNullOrWhiteSpace($ProfilePath)) { throw 'ProfilePath nao resolvido.' }
            $ini = Read-IniFile -Path $ProfilePath
            Log "$($script:ToolName) $($script:ToolVersion)"
            Log "RunId: $($script:RunId)"
            Log "Mode: $Mode"
            Log "Profile: $ProfilePath"
            Invoke-CompareBackendMode -Ini $ini
        }
        default          { throw "Modo invalido: $Mode" }
    }
    exit 0
}
catch {
    Log "ERRO FATAL: $($_.Exception.Message)" 'ERROR'
    try {
        if ([string]::IsNullOrWhiteSpace($script:CurrentRunRoot)) {
            $fallbackBase = 'C:\ECG\FieldKit\out'
            if ($script:CliOutDirProvided -and -not [string]::IsNullOrWhiteSpace($OutDir)) { $fallbackBase = $OutDir }
            $fallbackRoot = Join-Path $fallbackBase $script:RunId
            Ensure-Directory -Path $fallbackRoot
            $script:CurrentRunRoot = $fallbackRoot
        }

        if (-not (Test-Path -LiteralPath $script:CurrentRunRoot)) {
            New-Item -ItemType Directory -Path $script:CurrentRunRoot -Force | Out-Null
        }

        $fatalPath = Join-Path $script:CurrentRunRoot 'ECG_Fatal_Error.log'
        $fatalText = @(
            "ToolVersion=$($script:ToolVersion)"
            "CoreScriptPath=$($script:ScriptPath)"
            "ProfilePath=$ProfilePath"
            "TargetLabel=$TargetLabel"
            "RunId=$($script:RunId)"
            ($script:LogLines -join [Environment]::NewLine)
        ) -join [Environment]::NewLine
        Write-Utf8File -Path $fatalPath -Text $fatalText
    }
    catch {}
    exit 1
}
