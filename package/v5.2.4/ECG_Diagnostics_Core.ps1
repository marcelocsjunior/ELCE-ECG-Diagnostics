<#
.SYNOPSIS
    ECG Diagnostics Core - ferramenta unificada para diagnostico e correcao do ambiente ECG/BDE.
.DESCRIPTION
    Modos:
      Fix          : corrige NETDIR (HKLM/HKCU), IDAPI32.CFG + timeline + relatorio HTML/JSON
      Auto         : somente diagnostico (sem correcoes) + timeline + relatorio HTML/JSON
      Compare      : compara dois laudos JSON
      Rollback     : restaura backup .reg do BDE
      Monitor      : monitoramento prolongado com grafico, hipoteses e score
      CollectStatic: coleta estatica de informacoes (JSON)
.NOTES
    Versao: 5.2.4-unified-stable
#>

[CmdletBinding()]
param(
    [ValidateSet('Fix','Auto','Compare','Rollback','Monitor','CollectStatic')]
    [string]$Mode = 'Auto',

    [string]$ProfilePath = '',
    [string]$OutDir = 'C:\ECG\FieldKit\out',
    [string]$RollbackFile = '',
    [string]$CompareLeftReport = '',
    [string]$CompareRightReport = '',
    [switch]$OpenReport
)

$ErrorActionPreference = 'Stop'
$script:ToolName = 'ECG Diagnostics Core'
$script:ToolVersion = '5.2.4-unified-stable'
$script:ToolVersionShort = 'v5.2.4'
$script:HostName = $env:COMPUTERNAME
$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:LogLines = New-Object System.Collections.ArrayList
$script:RunId = (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:HostName
$script:CurrentRunRoot = ''
$script:CliOutDirProvided = $PSBoundParameters.ContainsKey('OutDir')

if ([string]::IsNullOrWhiteSpace($ProfilePath)) {
    try {
        $scriptDir = Split-Path -Parent $script:ScriptPath
        $candidateProfile = Join-Path $scriptDir 'ECG_FieldKit.ini'
        if (Test-Path $candidateProfile) { $ProfilePath = $candidateProfile }
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
    if (@('0','false','no','nao','não','off') -contains $raw) { return $false }
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
        'EXECUTANTE' { return 'Estação de exames' }
        'VIEWER'     { return 'Estação de visualização' }
        'HOST_XP'    { return 'Host XP legado' }
    }

    if ($computerUpper -match 'SRVVM1-FS01') { return 'Servidor de arquivos' }
    if ($computerUpper -match '(^|[-_])ECG([0-9A-Z_-]*$)') { return 'Estação de exames' }
    if ($computerUpper -match '(^|[-_])(CST|CON|VIEW)([0-9A-Z_-]*$)') { return 'Estação de visualização' }

    switch ($stationAlias) {
        'EXECUTANTE' { return 'Estação de exames' }
        'VIEWER'     { return 'Estação de visualização' }
        'HOST_XP'    { return 'Host XP legado' }
    }

    return 'Estação de trabalho'
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
    Write-Utf8File -Path $Path -Text ($lines -join [Environment]::NewLine)
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
        $changes += 'HW_CAMINHO_DB ajustado para o usuário atual'
    }
    catch {}
    if (Test-IsAdmin) {
        try {
            [Environment]::SetEnvironmentVariable('HW_CAMINHO_DB', $Value, 'Machine')
            $changes += 'HW_CAMINHO_DB ajustado em nível de máquina'
        }
        catch {}
    }
    return ,$changes
}

function Get-WmiSafe {
    param([string]$ClassName)
    try { return @(Get-WmiObject -Class $ClassName -ErrorAction Stop) } catch { return @() }
}

function Get-LocalShares { return @(Get-WmiSafe -ClassName 'Win32_Share' | Select-Object Name, Path) }
function Get-NetworkConnections { return @(Get-WmiSafe -ClassName 'Win32_NetworkConnection') }

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

function Resolve-EcgPaths {
    param($Profile)
    $officialDb = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedDbPath' -Default '\\192.168.1.57\Database')
    $officialNetDir = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedNetDir' -Default (Join-Path $officialDb 'NetDir'))
    $officialExe = Normalize-PathString (Get-ProfileString -Profile $Profile -Name 'ExpectedExePath' -Default 'C:\HW\ECG\ECGV6.exe')

    $effectiveDb = $officialDb
    $effectiveNetDir = $officialNetDir
    $dbSource = if (Test-IsUncPath $officialDb) { 'UNC oficial (INI)' } else { 'INI/valor oficial' }
    $netDirSource = if (Test-IsUncPath $officialNetDir) { 'UNC oficial (INI)' } else { 'INI/valor oficial' }

    $hwDb = Normalize-PathString $env:HW_CAMINHO_DB
    if ($hwDb -and (Test-Path $hwDb)) {
        if ($effectiveDb -ne $hwDb -and -not (Test-Path $effectiveDb)) {
            $effectiveDb = $hwDb
            $dbSource = 'Variável HW_CAMINHO_DB'
        }
        $candidateNet = Normalize-PathString (Join-Path $hwDb 'NetDir')
        if (-not (Test-Path $effectiveNetDir) -and $candidateNet -and (Test-Path $candidateNet)) {
            $effectiveNetDir = $candidateNet
            $netDirSource = 'Variável HW_CAMINHO_DB'
        }
    }

    $regCandidates = @(
        @{ Path = 'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT'; Source = 'Registro BDE (HKCU)' },
        @{ Path = 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT'; Source = 'Registro BDE (HKLM WOW6432Node)' },
        @{ Path = 'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT'; Source = 'Registro BDE (HKLM)' }
    )
    foreach ($candidate in $regCandidates) {
        $regNet = Normalize-PathString (Get-RegistryValueSafe -Path $candidate.Path -Name 'NETDIR')
        if ($regNet -and (Test-Path $regNet)) {
            $effectiveNetDir = $regNet
            $netDirSource = $candidate.Source
            break
        }
    }

    return [PSCustomObject]@{
        ExePath = $officialExe
        ExeAccessible = (Test-Path $officialExe)
        DatabasePath = $effectiveDb
        DatabasePathSource = $dbSource
        DatabaseAccessible = (Test-Path $effectiveDb)
        NetDirPath = $effectiveNetDir
        NetDirSource = $netDirSource
        NetDirAccessible = (Test-Path $effectiveNetDir)
        DesiredDbPath = $officialDb
        DesiredNetDir = $officialNetDir
        DesiredExePath = $officialExe
        DatabaseHost = Get-UncHostFromPath -Path $effectiveDb
        NetDirHost = Get-UncHostFromPath -Path $effectiveNetDir
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
    $pressure = if ($severity -ge 60) { 'Pressionado' } elseif ($severity -ge 30) { 'Atenção' } else { 'Estável' }

    return [PSCustomObject]@{
        Mode = 'Sem interação'
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
        $findings += 'Executável oficial do ECG não foi localizado no caminho padrão da ferramenta.'
        $scores.LOCAL += 4
        $hypothesisSupport += 'Executável oficial inacessível na estação durante a rodada.'
    }
    else {
        $discarded += 'Executável oficial localizado no caminho padrão.'
        $counterEvidence += 'Executável oficial acessível no caminho padrão.'
    }

    if (-not $Paths.DatabaseAccessible) {
        $findings += 'Banco do ECG inacessível no caminho efetivo da rodada.'
        $scores.SHARE += 5
        $hypothesisSupport += 'Banco inacessível no caminho efetivo durante a rodada.'
    }
    else {
        $discarded += 'Banco do ECG acessível no caminho efetivo da rodada.'
        $counterEvidence += 'Banco acessível no caminho efetivo nesta rodada.'
    }

    if (-not $Paths.NetDirAccessible) {
        $findings += 'NetDir inacessível no caminho efetivo da rodada.'
        $scores.SHARE += 5
        $hypothesisSupport += 'NetDir inacessível no caminho efetivo durante a rodada.'
    }
    else {
        $discarded += 'NetDir acessível no caminho efetivo da rodada.'
        $counterEvidence += 'NetDir acessível no caminho efetivo nesta rodada.'
    }

    $dbPathMatchesDesired = (Normalize-PathString $Paths.DatabasePath) -eq (Normalize-PathString $Paths.DesiredDbPath)
    if (-not $dbPathMatchesDesired) {
        $findings += 'A rodada não conseguiu operar com o caminho oficial do banco; foi necessário caminho alternativo.'
        $scores.LOCAL += 3
        $hypothesisSupport += 'Houve fallback de caminho para o banco, divergente do caminho oficial configurado.'
    }
    else {
        $discarded += 'Banco operando a partir do caminho oficial configurado nesta rodada.'
        $counterEvidence += 'Banco operando no caminho oficial configurado, sem fallback.'
    }

    $netDirMatchesDesired = (Normalize-PathString $Paths.NetDirPath) -eq (Normalize-PathString $Paths.DesiredNetDir)
    if (-not $netDirMatchesDesired) {
        $findings += 'NetDir efetivo divergente do caminho oficial configurado.'
        $scores.LOCAL += 2
        $hypothesisSupport += 'NetDir efetivo divergente do caminho oficial configurado.'
    }
    else {
        $discarded += 'NetDir operando no caminho oficial configurado nesta rodada.'
        $counterEvidence += 'NetDir operando no caminho oficial configurado.'
    }

    if ($dbUnavail -gt 0) {
        $findings += "Houve indisponibilidade do banco em $dbUnavail amostra(s) da rodada."
        $scores.SHARE += [math]::Min(4, $dbUnavail)
        $hypothesisSupport += "Banco instável em $dbUnavail amostra(s) da rodada."
    }
    else {
        $discarded += 'Sem indisponibilidade observada do banco durante a janela coletada.'
        $counterEvidence += 'Sem indisponibilidade de banco na janela observada.'
    }

    if ($netUnavail -gt 0) {
        $findings += "Houve indisponibilidade do NetDir em $netUnavail amostra(s) da rodada."
        $scores.SHARE += [math]::Min(4, $netUnavail)
        $hypothesisSupport += "NetDir instável em $netUnavail amostra(s) da rodada."
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
            $findings += "Compartilhamento acessível, porém com latência elevada na rodada (DB p95: $p95DbProbeMs ms | NetDir p95: $p95NetProbeMs ms)."
            $hypothesisSupport += 'Caminho compartilhado acessível, porém lento na janela observada.'
        }
        else {
            $scores.SHARE += 2
            $inconclusive += "Houve aumento de latência no compartilhamento sem indisponibilidade (DB p95: $p95DbProbeMs ms | NetDir p95: $p95NetProbeMs ms)."
        }
    }
    elseif (($null -ne $avgDbProbeMs) -or ($null -ne $avgNetProbeMs)) {
        $counterEvidence += "Latência de acesso ao compartilhamento em faixa saudável (DB média: $avgDbProbeMs ms | NetDir média: $avgNetProbeMs ms)."
    }

    $lockContentionStrong = $false
    if ($null -ne $peakLocks -and $peakLocks -ge 1) {
        if ($nominalLockBaseline) {
            $discarded += "Baseline de lock/controle compatível com operação compartilhada observado no NetDir (pico estável: $peakLocks)."
            $counterEvidence += "Locks em baseline nominal para ambiente compartilhado (pico estável: $peakLocks), sem indício adicional de incidente."
        }
        elseif ($peakLocks -ge 4 -or ($peakLocks -ge 3 -and ($dbUnavail -gt 0 -or $netUnavail -gt 0 -or $smbTimeout -gt 0)) -or ($null -ne $lockSpread -and $lockSpread -ge 2)) {
            $lockContentionStrong = $true
            $findings += "Atividade de lock/controle acima do baseline foi observada no NetDir durante a janela (pico: $peakLocks)."
            $scores.LOCK += [math]::Min(6, [int]$peakLocks)
            $hypothesisSupport += "Padrão de lock acima do baseline do NetDir (pico: $peakLocks)."
        }
        else {
            $inconclusive += "Locks/arquivos de controle foram vistos no NetDir (pico: $peakLocks), mas o sinal isolado ainda não sustenta contenção por si só."
        }
    }
    else {
        $discarded += 'Sem lock/controle relevante observado no NetDir durante a janela.'
        $counterEvidence += 'Sem lock/controle relevante no NetDir durante a janela observada.'
    }

    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 70) {
            $findings += "CPU média elevada na rodada ($avgCpu%)."
            $scores.SOFTWARE += 4
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $hypothesisSupport += "CPU média elevada com processo dominante no pico: $responsibleProcessSummary."
            }
            else {
                $hypothesisSupport += "CPU média elevada na rodada ($avgCpu%)."
            }
        }
        elseif ($avgCpu -lt 55) {
            $discarded += "CPU média da rodada sem pressão sustentada relevante ($avgCpu%)."
            $counterEvidence += "CPU média sem pressão sustentada relevante ($avgCpu%)."
        }
        else {
            $inconclusive += 'CPU média ficou em faixa intermediária; o sinal isolado não fecha hipótese por si só.'
        }
    }

    if ($null -ne $avgEcgCpu) {
        if ($avgEcgCpu -ge 40) {
            $findings += "Processo ECGV6 com uso relevante de CPU na rodada (média: $avgEcgCpu% | pico: $peakEcgCpu%)."
            $scores.SOFTWARE += 3
            $hypothesisSupport += 'Uso elevado de CPU associado diretamente ao processo ECGV6.'
        }
        elseif ($avgEcgCpu -lt 15) {
            $counterEvidence += "Processo ECGV6 sem pressão sustentada relevante de CPU (média: $avgEcgCpu%)."
        }
    }

    if ($null -ne $peakCpu) {
        if ($peakCpu -ge 95 -and $cpuBurstSamples90Plus -ge 3) {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $findings += "Burst local de CPU relevante observado na estação (pico: $peakCpu%). Processo dominante no pico: $responsibleProcessSummary."
                $hypothesisSupport += "Burst de CPU local relevante com processo dominante identificado: $responsibleProcessSummary."
            }
            else {
                $findings += "Burst local de CPU relevante observado na estação (pico: $peakCpu%)."
                $hypothesisSupport += "Burst de CPU local relevante observado na estação (pico: $peakCpu%)."
            }
            $scores.SOFTWARE += 4
        }
        elseif ($peakCpu -ge 80 -and $cpuBurstSamples90Plus -lt 3) {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                $inconclusive += "Houve pico isolado de CPU ($peakCpu%), porém sem persistência suficiente para fechar hipótese local. Processo dominante no pico: $responsibleProcessSummary."
            }
            else {
                $inconclusive += "Houve pico isolado de CPU ($peakCpu%), porém sem persistência suficiente para fechar hipótese local."
            }
        }
    }

    if ($null -ne $peakDiskQueueLength) {
        if ($peakDiskQueueLength -ge 8) {
            $findings += "Fila de disco elevada observada na rodada (pico: $peakDiskQueueLength)."
            $scores.SOFTWARE += 2
            $hypothesisSupport += 'Sinal de pressão de disco local na estação.'
        }
        elseif ($peakDiskQueueLength -ge 3) {
            $inconclusive += "Fila de disco moderada observada na rodada (pico: $peakDiskQueueLength)."
        }
    }

    if ($smbTimeout -gt 0) {
        $findings += 'Parte das consultas SMB excedeu timeout; leitura de compartilhamento deve ser interpretada com cautela.'
        $scores.SHARE += [math]::Min(3, $smbTimeout)
        $inconclusive += 'Consultas SMB com timeout reduzem a força de alguns descartes de compartilhamento.'
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
        'CRÍTICO'
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
    if ($status -eq 'NORMAL' -and $maxScore -eq 0) { $confidence = 'Média' }
    elseif ($maxScore -ge 8 -and $hypothesisSupport.Count -ge 2) { $confidence = 'Alta' }
    elseif ($maxScore -ge 4 -and $hypothesisSupport.Count -ge 1) { $confidence = 'Média' }

    $hypothesisMap = @{ LOCAL = 'Configuração local'; SHARE = 'Compartilhamento/acesso'; LOCK = 'Contenção/lock'; SOFTWARE = 'Software/arquivo'; OK = 'Ambiente íntegro' }
    $primaryHypothesis = $hypothesisMap[$primary]
    $impactScope = switch ($primary) { 'SHARE' { 'Sistema compartilhado' } 'LOCK' { 'Sistema compartilhado' } 'OK' { 'Sem impacto relevante observado' } default { 'Somente este computador' } }

    $recommendedAction = switch ($primary) {
        'SHARE' { 'Validar latência do compartilhamento, tempo de resposta do UNC e permissões antes de atuar no software.' }
        'LOCAL' { 'Conferir configuração local do ECG/BDE, mapeamentos e aderência aos caminhos oficiais da aplicação.' }
        'LOCK' { 'Repetir a rodada durante o sintoma e revisar concorrência de acesso, locks e fluxo de gravação no NetDir.' }
        'SOFTWARE' {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessName)) {
                if ($responsibleProcessIsEcg) {
                    "Revisar comportamento local do ECG, priorizando o processo $responsibleProcessName e correlacionando com o momento do pico de CPU."
                }
                else {
                    "Revisar a estação e o processo dominante no pico de CPU ($responsibleProcessName), verificando antivírus/EDR, tarefas concorrentes e interferência externa ao ECG."
                }
            }
            else {
                'Revisar saúde do software do ECG e comportamento local da estação, priorizando processo do ECG, antivírus/EDR e competidores de CPU.'
            }
        }
        default {
            if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
                "Sem ação corretiva imediata. Repetir a coleta apenas durante o sintoma. Na rodada atual, o processo dominante no pico foi: $responsibleProcessSummary."
            }
            else {
                'Sem ação corretiva imediata. Repetir a coleta apenas durante o sintoma percebido.'
            }
        }
    }

    $summaryPhrase = "A rodada indica condição $($status.ToLowerInvariant()) com hipótese principal em $($primaryHypothesis.ToLowerInvariant())."
    $probablePerception = switch ($status) {
        'CRÍTICO' { 'Usuário tende a perceber falha evidente, demora anormal ou impossibilidade prática de continuar o fluxo do ECG.' }
        'LENTO' { 'Usuário tende a perceber lentidão real ou resposta inconsistente durante a janela observada.' }
        'NORMAL' { 'A rodada atual não reuniu evidência de degradação ativa no backend compartilhado.' }
        default { 'Usuário pode ter percebido oscilação anterior, mas a rodada atual não reuniu evidência forte de degradação ativa.' }
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
        [string]$FallbackText = 'Sem evidência adicional.'
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
        $backendStatus = 'CRÍTICO'
        $backendLine = 'Há indisponibilidade do caminho efetivo do banco ou do NetDir na estação.'
    }
    elseif ($PassiveBenchmark.DatabaseUnavailableSamples -gt 0 -or $PassiveBenchmark.NetDirUnavailableSamples -gt 0 -or $PassiveBenchmark.SmbTimeoutSamples -gt 0) {
        $backendStatus = 'ATENÇÃO'
        $backendLine = 'Houve oscilação de backend/SMB durante a janela; validar compartilhamento antes de atuar no software.'
    }
    elseif (($null -ne $PassiveBenchmark.P95DatabaseProbeMs -and $PassiveBenchmark.P95DatabaseProbeMs -ge 100) -or ($null -ne $PassiveBenchmark.P95NetDirProbeMs -and $PassiveBenchmark.P95NetDirProbeMs -ge 100)) {
        $backendStatus = 'ATENÇÃO'
        $backendLine = "Compartilhamento acessível, porém com latência elevada (DB p95: $($PassiveBenchmark.P95DatabaseProbeMs) ms | NetDir p95: $($PassiveBenchmark.P95NetDirProbeMs) ms)."
    }

    $responsibleProcessSummary = [string]$PassiveBenchmark.PeakCpuResponsibleSummary

    $localStatus = 'OK'
    $localLine = 'Sem pressão local sustentada relevante nesta rodada.'
    if (-not $Paths.ExeAccessible) {
        $localStatus = 'ATENÇÃO'
        $localLine = 'Executável oficial não foi localizado no caminho esperado nesta estação.'
    }
    elseif (($null -ne $PassiveBenchmark.PeakCpuPercent -and $PassiveBenchmark.PeakCpuPercent -ge 95 -and $PassiveBenchmark.CpuBurstSamples90Plus -ge 3) -or ($null -ne $PassiveBenchmark.AverageCpuPercent -and $PassiveBenchmark.AverageCpuPercent -ge 70) -or ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent -and $PassiveBenchmark.PeakEcgProcessCpuPercent -ge 80)) {
        $localStatus = 'ATENÇÃO'
        if (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
            $localLine = "Pressão local observada. Processo dominante no pico: $responsibleProcessSummary."
        }
        elseif ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent -and $PassiveBenchmark.PeakEcgProcessCpuPercent -ge 80) {
            $localLine = "ECGV6 com uso elevado de CPU na rodada (pico: $($PassiveBenchmark.PeakEcgProcessCpuPercent)%)."
        }
        elseif ($PassiveBenchmark.CpuBurstSamples90Plus -ge 3 -and $null -ne $PassiveBenchmark.PeakCpuPercent) {
            $localLine = "Burst local de CPU observado (pico: $($PassiveBenchmark.PeakCpuPercent)% em $($PassiveBenchmark.CpuBurstSamples90Plus) amostra(s))."
        }
        elseif ($null -ne $PassiveBenchmark.AverageCpuPercent) {
            $localLine = "CPU média local em faixa elevada nesta rodada ($($PassiveBenchmark.AverageCpuPercent)%)."
        }
    }
    elseif (-not [string]::IsNullOrWhiteSpace($responsibleProcessSummary)) {
        $localLine = "Sem pressão local sustentada relevante nesta rodada. Processo dominante no pico isolado: $responsibleProcessSummary."
    }

    $lockLine = if ($PassiveBenchmark.NominalLockBaselineLikely) {
        "Locks em baseline nominal para ambiente compartilhado (pico estável: $($PassiveBenchmark.PeakLockFileCount))."
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
    param($Analysis, $MachineInfo, $Paths, $Timeline, $PassiveBenchmark, $StagePriority, $SymptomCode)
    $statusClass = switch ($Analysis.StatusCode) {
        'NORMAL' { 'normal' }
        'LENTO' { 'lento' }
        'CRÍTICO' { 'critico' }
        default { 'inconclusivo' }
    }

    $decision = Get-OperationalDecisionModel -Analysis $Analysis -Paths $Paths -PassiveBenchmark $PassiveBenchmark

    $topSupport = Select-TopTextItems -Items $Analysis.HypothesisSupport -Max 3 -FallbackText $(if ($Analysis.StatusCode -eq 'NORMAL') { 'Nenhum indício relevante de falha ativa nesta rodada.' } else { 'Sem evidência dominante suficiente nesta rodada.' })
    $topDiscard = Select-TopTextItems -Items $Analysis.WhatDidNotIndicateFailure -Max 5 -FallbackText 'Sem descarte adicional.'
    $topInconclusive = Select-TopTextItems -Items $Analysis.InconclusivePoints -Max 3 -FallbackText 'Sem ponto inconclusivo adicional.'
    $supportItems = Convert-TextItemsToHtmlList -Items $topSupport
    $discardItems = Convert-TextItemsToHtmlList -Items $topDiscard
    $inconclusiveItems = Convert-TextItemsToHtmlList -Items $topInconclusive

    $secondaryRows = if ($Analysis.SecondaryHypotheses.Count) {
        ($Analysis.SecondaryHypotheses | ForEach-Object {
            "<tr><td>$([System.Net.WebUtility]::HtmlEncode($_.Name))</td><td>$($_.Score)</td><td>Hipótese secundária com evidência parcial.</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="3">Sem hipótese secundária relevante.</td></tr>'
    }

    $relevantSamples = @(Get-RelevantTimelineSamples -Timeline $Timeline)
    $relevantTimelineRows = if ($relevantSamples.Count -gt 0) {
        ($relevantSamples | ForEach-Object {
            $dominantProcess = if ($_.TopCpuProcesses -and @($_.TopCpuProcesses).Count -gt 0) { Get-CpuResponsibleProcessSummary -ProcessInfo @($_.TopCpuProcesses)[0] -IncludeOwner:$false -IncludeCpu:$true } else { 'N/D' }
            "<tr><td>$($_.Timestamp)</td><td>$($_.CpuPercent)</td><td>$(if ($null -ne $_.EcgProcessCpuPercent) { $_.EcgProcessCpuPercent } else { 'N/D' })</td><td>$($_.LockFileCount)</td><td>$(if ($null -ne $_.DatabaseProbeMs) { $_.DatabaseProbeMs } else { 'N/D' })</td><td>$(if ($null -ne $_.NetDirProbeMs) { $_.NetDirProbeMs } else { 'N/D' })</td><td>$(if ($null -ne $_.PhysicalDiskQueueLength) { $_.PhysicalDiskQueueLength } else { 'N/D' })</td><td>$(if ($null -ne $_.NetworkBytesTotalPerSec) { $_.NetworkBytesTotalPerSec } else { 'N/D' })</td><td>$(if ($null -ne $_.SmbConnectionCount) { $_.SmbConnectionCount } else { 'N/A' })</td><td>$(if ($_.DatabaseAccessible) { 'Sim' } else { 'Não' })</td><td>$(if ($_.NetDirAccessible) { 'Sim' } else { 'Não' })</td><td>$(if ($_.SmbQueryTimedOut) { 'Sim' } else { 'Não' })</td><td>$([System.Net.WebUtility]::HtmlEncode($dominantProcess))</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="13">Sem eventos relevantes selecionados nesta rodada.</td></tr>'
    }

    $processCaptureRows = @(Get-ProcessCaptureRows -Timeline $Timeline)
    $processCaptureTableRows = if ($processCaptureRows.Count -gt 0) {
        ($processCaptureRows | ForEach-Object {
            "<tr><td>$($_.Timestamp)</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.ProcessName))</td><td>$($_.ProcessId)</td><td>$($_.CpuPercent)</td><td>$($_.ProcessCpuPercent)</td><td>$([System.Net.WebUtility]::HtmlEncode([string]$_.Owner))</td><td>$([System.Net.WebUtility]::HtmlEncode($(if ($_.IsEcgRelated) { 'Sim' } else { 'Não' })))</td></tr>"
        }) -join ''
    }
    else {
        '<tr><td colspan="7">Nenhuma captura de processo dominante foi necessária nesta rodada.</td></tr>'
    }

    $chartDef = Get-ChartDefinition -Timeline $Timeline
    $chartJson = $chartDef | ConvertTo-Json -Depth 10 -Compress
    $responsibleCpuText = if (-not [string]::IsNullOrWhiteSpace([string]$PassiveBenchmark.PeakCpuResponsibleSummary)) { [string]$PassiveBenchmark.PeakCpuResponsibleSummary } else { 'N/D' }

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
<div class="muted">Versão $script:ToolVersion</div>
<div class="grid-3">
<div class="kv"><strong>Máquina</strong>$($MachineInfo.ComputerName) — $($MachineInfo.MachineType)</div>
<div class="kv"><strong>Confiança</strong>$($Analysis.Confidence)</div>
<div class="kv"><strong>Observação</strong>$($Timeline.ObservationMinutes) minuto(s)</div>
<div class="kv"><strong>Hipótese principal</strong>$($Analysis.PrimaryHypothesis)</div>
<div class="kv"><strong>Impacto</strong>$($Analysis.ImpactScope)</div>
<div class="kv"><strong>Próxima ação</strong>$($Analysis.RecommendedAction)</div>
</div>
</div>

<div class="card">
<h2>Decisão operacional</h2>
<div class="grid">
<div class="kv"><strong>Backend compartilhado</strong>$($decision.BackendStatus)<br>$($decision.BackendLine)</div>
<div class="kv"><strong>Estação local</strong>$($decision.LocalStatus)<br>$($decision.LocalLine)</div>
</div>
<p><strong>Leitura operacional:</strong> $($Analysis.ProbablePerception)</p>
<p><strong>Leitura de lock:</strong> $($decision.LockLine)</p>
</div>

<div class="grid">
<div class="card"><h2>Evidências principais</h2><ul>$supportItems</ul></div>
<div class="card"><h2>Descartes relevantes</h2><ul>$discardItems</ul></div>
</div>

<div class="grid">
<div class="card"><h2>Métricas-chave</h2><div class="grid">
<div class="kv"><strong>Pressão</strong>$($PassiveBenchmark.PressureLabel)</div>
<div class="kv"><strong>Severidade</strong>$($PassiveBenchmark.SeverityScore)/100</div>
<div class="kv"><strong>CPU média / pico</strong>$($PassiveBenchmark.AverageCpuPercent)% / $($PassiveBenchmark.PeakCpuPercent)%</div>
<div class="kv"><strong>ECGV6 média / pico</strong>$(if ($null -ne $PassiveBenchmark.AverageEcgProcessCpuPercent) { $PassiveBenchmark.AverageEcgProcessCpuPercent } else { 'N/D' })% / $(if ($null -ne $PassiveBenchmark.PeakEcgProcessCpuPercent) { $PassiveBenchmark.PeakEcgProcessCpuPercent } else { 'N/D' })%</div>
<div class="kv"><strong>Locks mín/méd/pico</strong>$($PassiveBenchmark.MinimumLockFileCount) / $($PassiveBenchmark.AverageLockFileCount) / $($PassiveBenchmark.PeakLockFileCount)</div>
<div class="kv"><strong>DB ms média / p95</strong>$(if ($null -ne $PassiveBenchmark.AverageDatabaseProbeMs) { $PassiveBenchmark.AverageDatabaseProbeMs } else { 'N/D' }) / $(if ($null -ne $PassiveBenchmark.P95DatabaseProbeMs) { $PassiveBenchmark.P95DatabaseProbeMs } else { 'N/D' })</div>
<div class="kv"><strong>NetDir ms média / p95</strong>$(if ($null -ne $PassiveBenchmark.AverageNetDirProbeMs) { $PassiveBenchmark.AverageNetDirProbeMs } else { 'N/D' }) / $(if ($null -ne $PassiveBenchmark.P95NetDirProbeMs) { $PassiveBenchmark.P95NetDirProbeMs } else { 'N/D' })</div>
<div class="kv"><strong>DB / NetDir indisponível</strong>$($PassiveBenchmark.DatabaseUnavailableSamples) / $($PassiveBenchmark.NetDirUnavailableSamples)</div>
<div class="kv"><strong>Timeout SMB</strong>$($PassiveBenchmark.SmbTimeoutSamples)</div>
<div class="kv"><strong>Responsável CPU</strong>$([System.Net.WebUtility]::HtmlEncode($responsibleCpuText))</div>
<div class="kv"><strong>Capturas de processo</strong>$($PassiveBenchmark.SamplesWithProcessCapture)</div>
<div class="kv"><strong>Disk Queue pico</strong>$(if ($null -ne $PassiveBenchmark.PeakDiskQueueLength) { $PassiveBenchmark.PeakDiskQueueLength } else { 'N/D' })</div>
<div class="kv"><strong>Network Bytes pico</strong>$(if ($null -ne $PassiveBenchmark.PeakNetworkBytesTotalPerSec) { $PassiveBenchmark.PeakNetworkBytesTotalPerSec } else { 'N/D' })</div>
</div></div>
<div class="card"><h2>Pontos ainda inconclusivos</h2><ul>$inconclusiveItems</ul></div>
</div>

<details><summary>Detalhes técnicos</summary><div style="margin-top:16px;">
<div class="card"><h3>Contexto e paths</h3><table><tr><th>Campo</th><th>Valor</th></tr><tr><td>Executável oficial</td><td>$($Paths.ExePath) (acessível: $($Paths.ExeAccessible))</td></tr><tr><td>Banco efetivo</td><td>$($Paths.DatabasePath) ($($Paths.DatabasePathSource)) - acessível: $($Paths.DatabaseAccessible)</td></tr><tr><td>NetDir efetivo</td><td>$($Paths.NetDirPath) ($($Paths.NetDirSource)) - acessível: $($Paths.NetDirAccessible)</td></tr><tr><td>Locks baseline nominal</td><td>$($PassiveBenchmark.NominalLockBaselineLikely)</td></tr><tr><td>Processo dominante no pico de CPU</td><td>$([System.Net.WebUtility]::HtmlEncode($responsibleCpuText))</td></tr></table></div>
<div class="card"><h3>Eventos relevantes da timeline</h3><table><tr><th>Hora</th><th>CPU%</th><th>ECGV6%</th><th>Locks</th><th>DB ms</th><th>NetDir ms</th><th>DiskQ</th><th>Net Bytes/s</th><th>Conexões SMB</th><th>DB OK</th><th>NetDir OK</th><th>Timeout SMB</th><th>Processo dominante</th></tr>$relevantTimelineRows</table></div>
<div class="card"><h3>Processos capturados em alta CPU</h3><table><tr><th>Hora</th><th>Processo</th><th>PID</th><th>CPU amostra</th><th>CPU processo</th><th>Owner</th><th>ECG?</th></tr>$processCaptureTableRows</table></div>
<div class="card"><h3>Gráfico temporal</h3><canvas id="timelineChart" width="1100" height="320"></canvas><div id="timelineLegend" class="legend"></div><p class="muted">Séries normalizadas por escala própria.</p></div>
<div class="card"><h3>Hipóteses secundárias</h3><table><tr><th>Hipótese</th><th>Score</th><th>Observação</th></tr>$secondaryRows</table></div>
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
chartData.series.forEach(function(s){var leg=document.createElement('div');leg.className='legend-item';var cb=document.createElement('span');cb.className='legend-color';cb.style.backgroundColor=s.color;var txt=document.createElement('span');txt.textContent=s.name+' (máx='+s.max+')';leg.appendChild(cb);leg.appendChild(txt);legend.appendChild(leg);});
})();
</script>
</body>
</html>
"@
}

function Build-JsonReport {
    param($Analysis, $MachineInfo, $Paths, $Timeline, $PassiveBenchmark, $Profile, [string]$ModeName, [string[]]$AppliedChanges = @(), [string]$RollbackFile = '')
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
        AppliedChanges = @($AppliedChanges)
        RollbackFile = $RollbackFile
        Logs = @($script:LogLines)
    }
}

function Get-CompareSourceFiles {
    param([string]$RootPath)
    if (-not (Test-Path $RootPath)) { throw "OutDir não encontrado: $RootPath" }
    $files = Get-ChildItem -Path $RootPath -Recurse -Filter 'ECG_Report.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending
    if ($files.Count -lt 2) { throw "Preciso de pelo menos 2 relatórios JSON em $RootPath" }
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
<p>Comparação de convergência entre dois laudos JSON.</p>
<p class="muted"><strong>Esquerda:</strong> $([System.Net.WebUtility]::HtmlEncode($LeftPath))<br><strong>Direita:</strong> $([System.Net.WebUtility]::HtmlEncode($RightPath))</p>
</div>
<div class="card">
<h2>Campos críticos</h2>
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
    Log "Modo $Mode - Iniciado" 'STEP'

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

    if ($ApplyFixes -and -not (Test-IsAdmin)) {
        throw 'Modo Fix requer execução elevada (Administrador) para alterar HKLM/IDAPI32.CFG com segurança.'
    }

    $machineInfo = Get-KnownMachineInfo -ComputerName $script:HostName -Profile $profile
    $runRoot = Join-Path $effectiveOutDir $script:RunId
    $script:CurrentRunRoot = $runRoot
    New-Item -ItemType Directory -Path $runRoot -Force | Out-Null

    $appliedChanges = @()
    $rollbackFile = ''

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
                $appliedChanges += "NETDIR já estava aderente em $rp"
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
                $appliedChanges += "IDAPI32.CFG já aderente em $idapi"
            }
        }
    }

    $paths = Resolve-EcgPaths -Profile $profile
    $timeline = Collect-Timeline -MachineType $machineInfo.MachineType -Paths $paths -Minutes $observeMinutes -IntervalSeconds $sampleInterval -ProcessCaptureThresholdPercent $processCaptureThreshold -TopProcessCount $topProcessCount -EnableLatencyMetrics $enableLatencyMetrics -EnableEcgProcessMetrics $enableEcgProcessMetrics -EnableDiskMetrics $enableDiskMetrics -EnableNetworkMetrics $enableNetworkMetrics
    $benchmark = Build-PassiveBenchmark -Timeline $timeline -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'
    $analysis = Build-AnalysisModel -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'

    $html = Build-HtmlReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -StagePriority 'ABRIR_EXAME' -SymptomCode 'LENTIDAO_TRAVAMENTO'
    $modeLabel = if ($ApplyFixes) { 'Fix' } else { 'Auto' }
    $jsonData = Build-JsonReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $benchmark -Profile $profile -ModeName $modeLabel -AppliedChanges $appliedChanges -RollbackFile $rollbackFile

    $htmlPath = Join-Path $runRoot 'ECG_Report.html'
    $jsonPath = Join-Path $runRoot 'ECG_Report.json'
    Write-Utf8File -Path $htmlPath -Text $html
    Write-Utf8File -Path $jsonPath -Text ($jsonData | ConvertTo-Json -Depth 10)

    if ($OpenReport) { Start-Process $htmlPath }
    Log "Relatório HTML salvo em $htmlPath" 'STEP'
    Log "Relatório JSON salvo em $jsonPath" 'STEP'
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

    if (-not (Test-Path $left)) { throw "Relatório esquerdo não encontrado: $left" }
    if (-not (Test-Path $right)) { throw "Relatório direito não encontrado: $right" }

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
    Log "Comparação concluída: $status" 'STEP'
    Log "Relatório de comparação salvo em $outHtml" 'STEP'
}

function Invoke-RollbackMode {
    if (-not $RollbackFile) { throw 'Parâmetro -RollbackFile obrigatório' }
    if (-not (Test-Path $RollbackFile)) { throw "Arquivo $RollbackFile não encontrado" }
    if (-not (Test-IsAdmin)) { throw 'Rollback requer privilégios administrativos' }
    $script:CurrentRunRoot = Split-Path -Parent $RollbackFile
    $process = Start-Process -FilePath 'reg.exe' -ArgumentList @('import', $RollbackFile) -Wait -NoNewWindow -PassThru
    if ($process.ExitCode -ne 0) { throw "reg import retornou código $($process.ExitCode)" }
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
    Log "Relatório de monitoramento salvo em $htmlPath" 'STEP'
    Log "Relatório JSON de monitoramento salvo em $jsonPath" 'STEP'
}

function Invoke-CollectStaticMode {
    Log 'Coletando informações estáticas' 'STEP'
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
    Log "Coleta estática concluída: $jsonFile" 'STEP'
}

try {
    switch ($Mode) {
        'Fix'           { Invoke-FixAutoMode -ApplyFixes $true }
        'Auto'          { Invoke-FixAutoMode -ApplyFixes $false }
        'Compare'       { Invoke-CompareMode }
        'Rollback'      { Invoke-RollbackMode }
        'Monitor'       { Invoke-MonitorMode }
        'CollectStatic' { Invoke-CollectStaticMode }
        default         { throw "Modo inválido: $Mode" }
    }
    exit 0
}
catch {
    Log "ERRO FATAL: $($_.Exception.Message)" 'ERROR'
    try {
        if (-not [string]::IsNullOrWhiteSpace($script:CurrentRunRoot)) {
            if (-not (Test-Path $script:CurrentRunRoot)) { New-Item -ItemType Directory -Path $script:CurrentRunRoot -Force | Out-Null }
            $fatalPath = Join-Path $script:CurrentRunRoot 'ECG_Fatal_Error.log'
            $fatalText = @(
                "ToolVersion=$($script:ToolVersion)"
                "CoreScriptPath=$($script:ScriptPath)"
                "ProfilePath=$ProfilePath"
                "RunId=$($script:RunId)"
                ($script:LogLines -join [Environment]::NewLine)
            ) -join [Environment]::NewLine
            Write-Utf8File -Path $fatalPath -Text $fatalText
        }
    }
    catch {}
    exit 1
}
