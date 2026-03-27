<#
.SYNOPSIS
    ELCE ECG Diagnostics - Core reescrito com foco em laudo HTML.
.DESCRIPTION
    Ferramenta read-only para Windows PowerShell 5.1.
    Escopo desta versão:
      - execução padrão sem interação
      - índice observacional da rodada (passivo)
      - coleta temporal de 10 minutos (padrão)
      - geração de HTML principal
      - geração de context.json, benchmark.json, timeline.json e analysis.json
      - summary TXT/JSON em fail-soft
#>

[CmdletBinding()]
param(
    [ValidateSet('ABRIR_EXAME','SALVAR_FINALIZAR','GERAL')]
    [string]$StagePriority = 'ABRIR_EXAME',

    [ValidateSet('LENTIDAO_TRAVAMENTO','INCONCLUSIVO')]
    [string]$SymptomCode = 'LENTIDAO_TRAVAMENTO',

    [string]$SymptomText = '',

    [ValidateRange(1, 120)]
    [int]$ObservationMinutes = 10,

    [ValidateRange(5, 300)]
    [int]$SampleIntervalSeconds = 20,

    [string]$ParameterFile = '',

    [switch]$OpenReportOnSuccess
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ToolName = 'ELCE ECG Diagnostics'
$ToolVersion = '3.2-html-first-hotfix2'
$ToolRoot = 'C:\ECG\Tool'
$OutputRoot = 'C:\ECG\Output'
$RunsRoot = Join-Path $OutputRoot 'Runs'
$LatestRoot = Join-Path $OutputRoot 'Latest'

$OfficialDbPath = '\\SRVVM1-FS01\FS\ECG\HW\Database'
$OfficialNetDir = '\\SRVVM1-FS01\FS\ECG\HW\Database\Netdir'
$FallbackDbPath = 'P:\ECG\HW\Database'
$OfficialExePath = 'C:\HW\ECG\ECGV6.exe'
$FileServerHost = 'SRVVM1-FS01'

function Convert-ToSafeString {
    param($Value)

    if ($null -eq $Value) {
        return ''
    }

    if ($Value -is [System.Array]) {
        $items = New-Object System.Collections.Generic.List[string]
        foreach ($item in $Value) {
            if ($null -ne $item) {
                $items.Add([string]$item)
            }
        }
        return ($items -join [Environment]::NewLine)
    }

    return [string]$Value
}

function Write-Utf8NoBomFile {
    param(
        [string]$Path,
        $Content
    )

    $text = Convert-ToSafeString $Content
    $directory = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($Path, $text, $utf8NoBom)
}

function Append-Utf8NoBomFile {
    param(
        [string]$Path,
        $Content
    )

    $text = Convert-ToSafeString $Content
    $directory = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($directory) -and -not (Test-Path $directory)) {
        New-Item -Path $directory -ItemType Directory -Force | Out-Null
    }

    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::AppendAllText($Path, $text, $utf8NoBom)
}

function Save-JsonFile {
    param(
        [string]$Path,
        $Object,
        [int]$Depth = 8
    )

    $json = $null
    if ($null -eq $Object) {
        $json = 'null'
    }
    else {
        $json = $Object | ConvertTo-Json -Depth $Depth
    }
    Write-Utf8NoBomFile -Path $Path -Content ([string]$json)
}

function Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )

    $line = '{0} [{1}] {2}' -f (Get-Date -Format 'HH:mm:ss'), $Level, $Message
    Append-Utf8NoBomFile -Path $script:LogFile -Content ($line + [Environment]::NewLine)
    Write-Host $line
}

function HtmlEncode {
    param([string]$Value)
    if ($null -eq $Value) {
        return ''
    }
    return [System.Net.WebUtility]::HtmlEncode([string]$Value)
}

function Get-StageLabel {
    param([string]$Code)

    switch ($Code) {
        'ABRIR_EXAME' { return 'Abrir exame' }
        'SALVAR_FINALIZAR' { return 'Salvar/finalizar exame' }
        default { return 'Sem etapa prioritária definida' }
    }
}

function Get-SymptomLabel {
    param(
        [string]$Code,
        [string]$FreeText
    )

    switch ($Code) {
        'LENTIDAO_TRAVAMENTO' {
            if ([string]::IsNullOrWhiteSpace($FreeText)) {
                return 'Lentidão / travamentos'
            }
            return 'Lentidão / travamentos — ' + $FreeText.Trim()
        }
        default {
            if ([string]::IsNullOrWhiteSpace($FreeText)) {
                return 'Ainda sem sintoma informado'
            }
            return $FreeText.Trim()
        }
    }
}

function Get-RegistryValueSafe {
    param(
        [string]$Path,
        [string]$Name
    )

    try {
        $item = Get-ItemProperty -Path $Path -ErrorAction Stop
        return $item.$Name
    }
    catch {
        return $null
    }
}

function Get-KnownMachineInfo {
    param([string]$ComputerName)

    $machineType = 'Tipo ainda não definido'
    $expectedUser = ''

    switch ($ComputerName.ToUpperInvariant()) {
        'SRVVM1-FS01' {
            $machineType = 'Servidor de arquivos'
        }
        'ELCUN1-ECG' {
            $machineType = 'Estação de exames'
            $expectedUser = 'elce\ecg.un1'
        }
        'ELCUN1-CST2' {
            $machineType = 'Estação de visualização'
            $expectedUser = 'elce\ewaldo.bayao'
        }
    }

    $executedBy = ''
    try {
        $executedBy = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        $executedBy = [string]$env:USERNAME
    }

    $expectedUserMatch = $false
    if (-not [string]::IsNullOrWhiteSpace($expectedUser) -and -not [string]::IsNullOrWhiteSpace($executedBy)) {
        $expectedUserMatch = ([string]$expectedUser).ToLowerInvariant() -eq ([string]$executedBy).ToLowerInvariant()
    }

    return [PSCustomObject]@{
        ComputerName = $ComputerName
        MachineType = $machineType
        ExecutedBy = $executedBy
        ExpectedUser = $expectedUser
        ExpectedUserMatch = $expectedUserMatch
    }
}

function Resolve-EcgPaths {
    $effectiveDb = $OfficialDbPath
    $effectiveNetDir = $OfficialNetDir
    $dbSource = 'UNC direto'
    $netDirSource = 'UNC direto'
    $pathNotes = New-Object System.Collections.Generic.List[string]

    if (-not (Test-Path $effectiveDb)) {
        $regDb = $null
        foreach ($candidate in @(
            'HKCU:\Software\HeartWare\ECGV6\Geral',
            'HKLM:\Software\HeartWare\ECGV6\Geral',
            'HKLM:\Software\WOW6432Node\HeartWare\ECGV6\Geral'
        )) {
            $tmp = Get-RegistryValueSafe -Path $candidate -Name 'Caminho Database'
            if (-not [string]::IsNullOrWhiteSpace($tmp)) {
                $regDb = [string]$tmp
                break
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($regDb)) {
            $pathNotes.Add('Banco oficial indisponível; validando configuração local ECG/BDE.')
            if (Test-Path $regDb) {
                $effectiveDb = $regDb
                $dbSource = 'Configuração do ECG/BDE'
                $candidateNetDir = Join-Path $regDb 'Netdir'
                if (Test-Path $candidateNetDir) {
                    $effectiveNetDir = $candidateNetDir
                    $netDirSource = 'Configuração do ECG/BDE'
                }
            }
        }

        if (($dbSource -eq 'UNC direto') -and (Test-Path $FallbackDbPath)) {
            $effectiveDb = $FallbackDbPath
            $dbSource = 'Unidade mapeada'
            $fallbackNetDir = Join-Path $FallbackDbPath 'Netdir'
            if (Test-Path $fallbackNetDir) {
                $effectiveNetDir = $fallbackNetDir
                $netDirSource = 'Unidade mapeada'
            }
            $pathNotes.Add('Banco oficial indisponível; usando fallback por unidade mapeada.')
        }
    }

    return [PSCustomObject]@{
        ExePath = $OfficialExePath
        ExeAccessible = [bool](Test-Path $OfficialExePath)
        DatabasePath = $effectiveDb
        DatabasePathSource = $dbSource
        DatabaseAccessible = [bool](Test-Path $effectiveDb)
        NetDirPath = $effectiveNetDir
        NetDirSource = $netDirSource
        NetDirAccessible = [bool](Test-Path $effectiveNetDir)
        Notes = @($pathNotes)
    }
}

function Invoke-JobWithTimeout {
    param(
        [scriptblock]$ScriptBlock,
        [object[]]$ArgumentList = @(),
        [int]$TimeoutSeconds = 4
    )

    $job = $null
    try {
        $job = Start-Job -ScriptBlock $ScriptBlock -ArgumentList $ArgumentList
        $completed = Wait-Job -Job $job -Timeout $TimeoutSeconds
        if ($null -eq $completed) {
            try { Stop-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
            return [PSCustomObject]@{
                Completed = $false
                TimedOut = $true
                Output = @()
            }
        }

        $output = @()
        try {
            $output = @(Receive-Job -Job $job -ErrorAction SilentlyContinue)
        }
        catch {
            $output = @()
        }
        try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } catch {}

        return [PSCustomObject]@{
            Completed = $true
            TimedOut = $false
            Output = @($output)
        }
    }
    catch {
        if ($null -ne $job) {
            try { Remove-Job -Job $job -Force -ErrorAction SilentlyContinue | Out-Null } catch {}
        }
        return [PSCustomObject]@{
            Completed = $false
            TimedOut = $false
            Output = @()
        }
    }
}

function Get-CpuPercentSafe {
    try {
        $sample = (Get-Counter '\Processor(_Total)\% Processor Time' -ErrorAction Stop).CounterSamples[0].CookedValue
        return [math]::Round([double]$sample, 2)
    }
    catch {
        try {
            $cpu = (Get-CimInstance Win32_Processor -ErrorAction Stop | Measure-Object -Property LoadPercentage -Average).Average
            if ($null -ne $cpu) {
                return [math]::Round([double]$cpu, 2)
            }
        }
        catch {}
    }

    return $null
}

function Get-SmbMetricsSafe {
    param($MachineInfo, $Paths)

    $result = [ordered]@{
        SmbConnectionCount = $null
        SmbSessionCount = $null
        RelevantOpenFileCount = $null
        RelevantNetDirOpenFileCount = $null
        SmbQueryTimedOut = $false
    }

    if ($MachineInfo.MachineType -eq 'Servidor de arquivos') {
        try {
            $sessionCall = Invoke-JobWithTimeout -TimeoutSeconds 4 -ScriptBlock {
                if (Get-Command -Name Get-SmbSession -ErrorAction SilentlyContinue) {
                    @(Get-SmbSession -ErrorAction SilentlyContinue)
                }
            }

            if ($sessionCall.Completed) {
                $sessions = @($sessionCall.Output)
                $result.SmbSessionCount = $sessions.Count
            }
            elseif ($sessionCall.TimedOut) {
                $result.SmbQueryTimedOut = $true
            }
        }
        catch {}

        try {
            $openFileCall = Invoke-JobWithTimeout -TimeoutSeconds 4 -ScriptBlock {
                if (Get-Command -Name Get-SmbOpenFile -ErrorAction SilentlyContinue) {
                    @(Get-SmbOpenFile -ErrorAction SilentlyContinue)
                }
            }

            if ($openFileCall.Completed) {
                $openFiles = @($openFileCall.Output)
                if ($openFiles.Count -gt 0) {
                    $relevantOpen = @($openFiles | Where-Object { $_.Path -and ($_.Path -like '*\FS\ECG\HW\Database*') })
                    $relevantNet = @($openFiles | Where-Object { $_.Path -and ($_.Path -like '*\FS\ECG\HW\Database\Netdir*') })
                    $result.RelevantOpenFileCount = $relevantOpen.Count
                    $result.RelevantNetDirOpenFileCount = $relevantNet.Count
                }
                else {
                    $result.RelevantOpenFileCount = 0
                    $result.RelevantNetDirOpenFileCount = 0
                }
            }
            elseif ($openFileCall.TimedOut) {
                $result.SmbQueryTimedOut = $true
            }
        }
        catch {}
    }
    else {
        try {
            $connCall = Invoke-JobWithTimeout -TimeoutSeconds 4 -ScriptBlock {
                param($ServerName)
                if (Get-Command -Name Get-SmbConnection -ErrorAction SilentlyContinue) {
                    @(Get-SmbConnection -ErrorAction SilentlyContinue | Where-Object {
                        $_.ServerName -eq $ServerName -or $_.ServerName -eq 'SRVVM1-FS01'
                    })
                }
            } -ArgumentList @($FileServerHost)

            if ($connCall.Completed) {
                $connections = @($connCall.Output)
                $result.SmbConnectionCount = $connections.Count
            }
            elseif ($connCall.TimedOut) {
                $result.SmbQueryTimedOut = $true
            }
        }
        catch {}
    }

    return [PSCustomObject]$result
}

function Get-LockFileCountSafe {
    param([string]$NetDirPath)

    try {
        if (Test-Path $NetDirPath) {
            $lockFiles = @(Get-ChildItem -Path $NetDirPath -File -ErrorAction SilentlyContinue | Where-Object {
                $_.Name -match 'PDOXUSRS|\.LCK$|\.NET$'
            })
            return $lockFiles.Count
        }
    }
    catch {}

    return $null
}

function Get-TimelineSample {
    param($MachineInfo, $Paths, [int]$SampleIndex)

    $sampleTime = Get-Date
    $dbOk = [bool](Test-Path $Paths.DatabasePath)
    $netDirOk = [bool](Test-Path $Paths.NetDirPath)
    $exeOk = [bool](Test-Path $Paths.ExePath)
    $cpu = Get-CpuPercentSafe
    $lockCount = Get-LockFileCountSafe -NetDirPath $Paths.NetDirPath
    $smb = Get-SmbMetricsSafe -MachineInfo $MachineInfo -Paths $Paths

    return [PSCustomObject]@{
        SampleIndex = $SampleIndex
        Timestamp = $sampleTime.ToString('HH:mm:ss')
        SampleTimeIso = $sampleTime.ToString('s')
        CpuPercent = $cpu
        DatabaseAccessible = $dbOk
        NetDirAccessible = $netDirOk
        ExeAccessible = $exeOk
        LockFileCount = $lockCount
        SmbConnectionCount = $smb.SmbConnectionCount
        SmbSessionCount = $smb.SmbSessionCount
        RelevantOpenFileCount = $smb.RelevantOpenFileCount
        RelevantNetDirOpenFileCount = $smb.RelevantNetDirOpenFileCount
        SmbQueryTimedOut = [bool]$smb.SmbQueryTimedOut
    }
}

function Collect-Timeline {
    param(
        $MachineInfo,
        $Paths,
        [int]$Minutes,
        [int]$IntervalSeconds
    )

    $samples = @()
    $started = Get-Date
    $sampleTarget = [int][math]::Ceiling(($Minutes * 60) / $IntervalSeconds)
    if ($sampleTarget -lt 1) {
        $sampleTarget = 1
    }

    Log ('Coleta temporal iniciada. Janela=' + $Minutes + ' min | Intervalo=' + $IntervalSeconds + ' s | Meta de amostras=' + $sampleTarget) 'STEP'

    for ($i = 1; $i -le $sampleTarget; $i++) {
        $sample = Get-TimelineSample -MachineInfo $MachineInfo -Paths $Paths -SampleIndex $i
        $samples += $sample

        $cpuText = 'N/A'
        if ($null -ne $sample.CpuPercent) {
            $cpuText = ([string]$sample.CpuPercent) + '%'
        }

        $dbText = [string]$sample.DatabaseAccessible
        $netText = [string]$sample.NetDirAccessible
        $lockText = if ($null -eq $sample.LockFileCount) { 'N/A' } else { [string]$sample.LockFileCount }
        Log ('Amostra ' + $i + '/' + $sampleTarget + ' | Hora=' + $sample.Timestamp + ' | CPU=' + $cpuText + ' | Locks=' + $lockText + ' | DB=' + $dbText + ' | NetDir=' + $netText) 'INFO'

        if ($i -lt $sampleTarget) {
            Start-Sleep -Seconds $IntervalSeconds
        }
    }

    $ended = Get-Date

    return [PSCustomObject]@{
        StartedAt = $started.ToString('dd/MM/yyyy HH:mm:ss')
        EndedAt = $ended.ToString('dd/MM/yyyy HH:mm:ss')
        ObservationMinutes = $Minutes
        IntervalSeconds = $IntervalSeconds
        SampleCount = $samples.Count
        Samples = @($samples)
    }
}

function Get-MaxMetricValue {
    param(
        [array]$Samples,
        [string]$PropertyName
    )

    $values = New-Object System.Collections.Generic.List[double]
    foreach ($sample in @($Samples)) {
        if ($null -ne $sample -and $sample.PSObject.Properties.Name -contains $PropertyName) {
            $value = $sample.$PropertyName
            if ($null -ne $value -and "$value" -ne '') {
                try {
                    $values.Add([double]$value)
                }
                catch {}
            }
        }
    }

    if ($values.Count -eq 0) {
        return $null
    }

    return [math]::Round((($values | Measure-Object -Maximum).Maximum), 2)
}

function Get-AverageMetricValue {
    param(
        [array]$Samples,
        [string]$PropertyName
    )

    $values = New-Object System.Collections.Generic.List[double]
    foreach ($sample in @($Samples)) {
        if ($null -ne $sample -and $sample.PSObject.Properties.Name -contains $PropertyName) {
            $value = $sample.$PropertyName
            if ($null -ne $value -and "$value" -ne '') {
                try {
                    $values.Add([double]$value)
                }
                catch {}
            }
        }
    }

    if ($values.Count -eq 0) {
        return $null
    }

    return [math]::Round((($values | Measure-Object -Average).Average), 2)
}

function Get-BooleanFalseCount {
    param(
        [array]$Samples,
        [string]$PropertyName
    )

    $count = 0
    foreach ($sample in @($Samples)) {
        if ($null -ne $sample -and $sample.PSObject.Properties.Name -contains $PropertyName) {
            $value = $sample.$PropertyName
            if ($value -eq $false) {
                $count++
            }
        }
    }
    return $count
}

function Build-PassiveBenchmark {
    param(
        $Timeline,
        [string]$StagePriority,
        [string]$SymptomCode
    )

    $samples = @($Timeline.Samples)
    $sampleCount = $samples.Count
    if ($sampleCount -lt 1) {
        $sampleCount = 1
    }

    $avgCpu = Get-AverageMetricValue -Samples $samples -PropertyName 'CpuPercent'
    $maxCpu = Get-MaxMetricValue -Samples $samples -PropertyName 'CpuPercent'
    $peakLocks = Get-MaxMetricValue -Samples $samples -PropertyName 'LockFileCount'
    $peakSmbConnections = Get-MaxMetricValue -Samples $samples -PropertyName 'SmbConnectionCount'
    $peakSmbSessions = Get-MaxMetricValue -Samples $samples -PropertyName 'SmbSessionCount'
    $peakRelevantOpen = Get-MaxMetricValue -Samples $samples -PropertyName 'RelevantOpenFileCount'
    $peakRelevantNetDirOpen = Get-MaxMetricValue -Samples $samples -PropertyName 'RelevantNetDirOpenFileCount'

    $dbUnavailableSamples = Get-BooleanFalseCount -Samples $samples -PropertyName 'DatabaseAccessible'
    $netDirUnavailableSamples = Get-BooleanFalseCount -Samples $samples -PropertyName 'NetDirAccessible'
    $exeUnavailableSamples = Get-BooleanFalseCount -Samples $samples -PropertyName 'ExeAccessible'
    $smbTimeoutSamples = 0
    foreach ($sample in $samples) {
        if ($sample.SmbQueryTimedOut -eq $true) {
            $smbTimeoutSamples++
        }
    }

    $severityScore = 0
    if ($dbUnavailableSamples -gt 0) {
        $severityScore += [math]::Min(40, ($dbUnavailableSamples * 10))
    }
    if ($netDirUnavailableSamples -gt 0) {
        $severityScore += [math]::Min(35, ($netDirUnavailableSamples * 10))
    }
    if ($exeUnavailableSamples -gt 0) {
        $severityScore += 45
    }
    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 85) {
            $severityScore += 25
        }
        elseif ($avgCpu -ge 70) {
            $severityScore += 15
        }
        elseif ($avgCpu -ge 55) {
            $severityScore += 5
        }
    }
    if ($null -ne $maxCpu) {
        if ($maxCpu -ge 95) {
            $severityScore += 15
        }
        elseif ($maxCpu -ge 80) {
            $severityScore += 8
        }
    }
    if ($null -ne $peakLocks) {
        if ($peakLocks -ge 5) {
            $severityScore += 18
        }
        elseif ($peakLocks -ge 1) {
            $severityScore += 10
        }
    }
    if ($smbTimeoutSamples -gt 0) {
        $severityScore += [math]::Min(10, ($smbTimeoutSamples * 3))
    }

    if ($severityScore -gt 100) {
        $severityScore = 100
    }

    $pressureLabel = 'Estável'
    if ($severityScore -ge 60) {
        $pressureLabel = 'Pressionado'
    }
    elseif ($severityScore -ge 30) {
        $pressureLabel = 'Atenção'
    }

    return [PSCustomObject]@{
        BenchmarkType = 'PASSIVE_OBSERVATIONAL'
        StageCode = $StagePriority
        StageLabel = Get-StageLabel -Code $StagePriority
        SymptomCode = $SymptomCode
        Mode = 'Sem interação'
        StartedAt = $Timeline.StartedAt
        EndedAt = $Timeline.EndedAt
        ObservationMinutes = $Timeline.ObservationMinutes
        IntervalSeconds = $Timeline.IntervalSeconds
        SampleCount = $Timeline.SampleCount
        AverageCpuPercent = $avgCpu
        PeakCpuPercent = $maxCpu
        PeakLockFileCount = $peakLocks
        PeakSmbConnectionCount = $peakSmbConnections
        PeakSmbSessionCount = $peakSmbSessions
        PeakRelevantOpenFileCount = $peakRelevantOpen
        PeakRelevantNetDirOpenFileCount = $peakRelevantNetDirOpen
        DatabaseUnavailableSamples = $dbUnavailableSamples
        NetDirUnavailableSamples = $netDirUnavailableSamples
        ExeUnavailableSamples = $exeUnavailableSamples
        SmbTimeoutSamples = $smbTimeoutSamples
        SeverityScore = $severityScore
        PressureLabel = $pressureLabel
        SummaryPhrase = 'Índice observacional da rodada calculado de forma passiva, sem benchmark assistido dentro do ECG.'
    }
}

function Build-AnalysisModel {
    param(
        $MachineInfo,
        $Paths,
        $Timeline,
        $PassiveBenchmark
    )

    $findings = New-Object System.Collections.Generic.List[string]
    $discarded = New-Object System.Collections.Generic.List[string]
    $limitations = New-Object System.Collections.Generic.List[string]
    $signals = New-Object System.Collections.Generic.List[string]

    $avgCpu = $PassiveBenchmark.AverageCpuPercent
    $peakCpu = $PassiveBenchmark.PeakCpuPercent
    $peakLocks = $PassiveBenchmark.PeakLockFileCount
    $dbUnavailableSamples = [int]$PassiveBenchmark.DatabaseUnavailableSamples
    $netUnavailableSamples = [int]$PassiveBenchmark.NetDirUnavailableSamples
    $exeUnavailableSamples = [int]$PassiveBenchmark.ExeUnavailableSamples
    $smbTimeoutSamples = [int]$PassiveBenchmark.SmbTimeoutSamples
    $severityScore = [int]$PassiveBenchmark.SeverityScore
    $sampleCount = [int]$Timeline.SampleCount
    if ($sampleCount -lt 1) {
        $sampleCount = 1
    }

    $dbUnavailableRatio = [math]::Round(($dbUnavailableSamples / [double]$sampleCount), 2)
    $netUnavailableRatio = [math]::Round(($netUnavailableSamples / [double]$sampleCount), 2)

    if (-not $Paths.ExeAccessible) {
        $findings.Add('Executável oficial do ECG não foi localizado no caminho padrão da ferramenta.')
        $signals.Add('EXE_OFFICIAL_NOT_FOUND')
    }
    else {
        $discarded.Add('Executável oficial do ECG localizado no caminho padrão.')
    }

    if (-not $Paths.DatabaseAccessible) {
        $findings.Add('Banco do ECG inacessível no caminho efetivo da rodada.')
        $signals.Add('DATABASE_PATH_UNAVAILABLE')
    }
    else {
        $discarded.Add('Banco do ECG acessível no caminho efetivo da rodada.')
    }

    if (-not $Paths.NetDirAccessible) {
        $findings.Add('NetDir inacessível no caminho efetivo da rodada.')
        $signals.Add('NETDIR_PATH_UNAVAILABLE')
    }
    else {
        $discarded.Add('NetDir acessível no caminho efetivo da rodada.')
    }

    if ($Paths.DatabasePathSource -ne 'UNC direto') {
        $findings.Add('A rodada não conseguiu operar apenas com o UNC oficial do banco; foi necessário caminho alternativo.')
        $signals.Add('DATABASE_SOURCE_FALLBACK')
    }
    else {
        $discarded.Add('Banco operando a partir do UNC oficial nesta rodada.')
    }

    if ($dbUnavailableSamples -gt 0) {
        $findings.Add('Houve indisponibilidade do banco em ' + $dbUnavailableSamples + ' amostra(s) da rodada.')
        $signals.Add('DATABASE_ACCESS_INSTABILITY')
    }
    else {
        $discarded.Add('Sem indisponibilidade observada do banco durante a janela coletada.')
    }

    if ($netUnavailableSamples -gt 0) {
        $findings.Add('Houve indisponibilidade do NetDir em ' + $netUnavailableSamples + ' amostra(s) da rodada.')
        $signals.Add('NETDIR_ACCESS_INSTABILITY')
    }
    else {
        $discarded.Add('Sem indisponibilidade observada do NetDir durante a janela coletada.')
    }

    if ($null -ne $peakLocks -and $peakLocks -ge 1) {
        $findings.Add('Arquivos de lock/controle foram observados no NetDir durante a janela (pico: ' + $peakLocks + ').')
        $signals.Add('LOCK_ACTIVITY_OBSERVED')
    }
    else {
        $discarded.Add('Sem lock/controle relevante observado no NetDir durante a janela.')
    }

    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 70) {
            $findings.Add('CPU média da rodada em faixa elevada (' + $avgCpu + '%).')
            $signals.Add('CPU_AVERAGE_HIGH')
        }
        elseif ($avgCpu -lt 55) {
            $discarded.Add('CPU média da rodada sem pressão sustentada relevante (' + $avgCpu + '%).')
        }
    }

    if ($null -ne $peakCpu) {
        if ($peakCpu -ge 90) {
            $findings.Add('Pico de CPU relevante durante a rodada (' + $peakCpu + '%).')
            $signals.Add('CPU_PEAK_HIGH')
        }
        elseif ($peakCpu -lt 80) {
            $discarded.Add('Sem pico extremo de CPU durante a rodada (' + $peakCpu + '%).')
        }
    }

    if ($smbTimeoutSamples -gt 0) {
        $findings.Add('Parte das consultas SMB excedeu timeout; leitura de compartilhamento deve ser interpretada com cautela.')
        $signals.Add('SMB_TIMEOUTS')
    }
    else {
        $discarded.Add('Sem timeout observado nas consultas SMB executadas pela ferramenta.')
    }

    $limitations.Add('Esta versão não executa benchmark assistido dentro do ECG; usa índice observacional da rodada.')
    $limitations.Add('Esta versão não executa comparação com referência.')
    $limitations.Add('Esta versão não executa avaliação Defender/minifilter.')

    $probableCause = 'Ainda sem causa definida'
    if ((-not $Paths.DatabaseAccessible) -or (-not $Paths.NetDirAccessible) -or ($dbUnavailableRatio -ge 0.20) -or ($netUnavailableRatio -ge 0.20)) {
        $probableCause = 'Provável no compartilhamento/acesso'
    }
    elseif ((-not $Paths.ExeAccessible) -or ($Paths.DatabasePathSource -ne 'UNC direto')) {
        $probableCause = 'Provável na configuração local'
    }
    elseif ($null -ne $peakLocks -and $peakLocks -ge 1) {
        $probableCause = 'Provável por contenção/lock'
    }
    elseif ((($null -ne $avgCpu) -and ($avgCpu -ge 70)) -or (($null -ne $peakCpu) -and ($peakCpu -ge 90))) {
        $probableCause = 'Provável no software/arquivo'
    }

    $status = 'INCONCLUSIVO'
    if ((-not $Paths.ExeAccessible) -or ($dbUnavailableRatio -ge 0.30) -or ($netUnavailableRatio -ge 0.30)) {
        $status = 'CRÍTICO'
    }
    elseif (($severityScore -ge 30) -or (($null -ne $peakLocks) -and ($peakLocks -ge 1)) -or (($null -ne $avgCpu) -and ($avgCpu -ge 65))) {
        $status = 'LENTO'
    }
    elseif ($severityScore -eq 0 -and $Paths.ExeAccessible -and $Paths.DatabaseAccessible -and $Paths.NetDirAccessible) {
        $status = 'NORMAL'
    }

    $confidence = 'Baixa'
    if ($status -eq 'CRÍTICO' -and ($probableCause -ne 'Ainda sem causa definida')) {
        $confidence = 'Alta'
    }
    elseif ($status -eq 'LENTO' -and ($probableCause -ne 'Ainda sem causa definida')) {
        $confidence = 'Média'
    }
    elseif ($status -eq 'NORMAL') {
        $confidence = 'Média'
    }

    $impactScope = 'Ainda não foi possível definir'
    if ($MachineInfo.MachineType -eq 'Servidor de arquivos') {
        $impactScope = 'Sistema compartilhado'
    }
    elseif ($probableCause -eq 'Provável no compartilhamento/acesso' -or $probableCause -eq 'Provável por contenção/lock') {
        $impactScope = 'Sistema compartilhado'
    }
    elseif ($probableCause -eq 'Provável na configuração local' -or $probableCause -eq 'Provável no software/arquivo') {
        $impactScope = 'Somente este computador'
    }

    $probablePerception = 'Usuário percebe lentidão e travamentos intermitentes no ECG, principalmente ao abrir exame.'
    if ($status -eq 'NORMAL') {
        $probablePerception = 'Usuário pode ter percebido oscilação anterior, mas a rodada atual não reuniu evidência forte de degradação ativa.'
    }
    elseif ($status -eq 'CRÍTICO') {
        $probablePerception = 'Usuário tende a perceber falha evidente, demora anormal ou impossibilidade prática de continuar o fluxo do ECG.'
    }

    $recommendedAction = 'Executar nova rodada durante o sintoma e cruzar com observação do operador para consolidar a hipótese.'
    switch ($probableCause) {
        'Provável no compartilhamento/acesso' {
            $recommendedAction = 'Validar disponibilidade do caminho do banco/NetDir e revisar camada de compartilhamento antes de atuar no software.'
        }
        'Provável por contenção/lock' {
            $recommendedAction = 'Repetir a rodada durante o sintoma e revisar contenção/locks no NetDir e concorrência de acesso.'
        }
        'Provável na configuração local' {
            $recommendedAction = 'Conferir configuração local do ECG/BDE e garantir aderência aos caminhos oficiais da aplicação.'
        }
        'Provável no software/arquivo' {
            $recommendedAction = 'Revisar saúde do software do ECG e comportamento local da estação, priorizando logs e estado do aplicativo.'
        }
        default {
            if ($status -eq 'NORMAL') {
                $recommendedAction = 'Manter observação; a rodada atual não mostrou evidência técnica forte de degradação ativa.'
            }
        }
    }

    $summaryPhrase = 'A rodada não reuniu evidência suficiente para fechar causa com segurança.'
    if ($status -eq 'NORMAL') {
        $summaryPhrase = 'A rodada foi concluída sem evidência forte de degradação ativa nas camadas observadas pela ferramenta.'
    }
    elseif ($status -eq 'LENTO') {
        $summaryPhrase = 'A rodada indica lentidão operacional com maior suspeita em ' + ($probableCause -replace '^Provável ', '').ToLowerInvariant() + '.'
    }
    elseif ($status -eq 'CRÍTICO') {
        $summaryPhrase = 'A rodada indica condição crítica com impacto operacional imediato, exigindo ação prioritária.'
    }

    if ($findings.Count -eq 0) {
        $findings.Add('A rodada não encontrou evidência forte suficiente para afirmar uma causa provável acima das demais.')
    }
    if ($discarded.Count -eq 0) {
        $discarded.Add('A rodada ainda não produziu descartes técnicos fortes.')
    }

    $statusDisplay = $status.Substring(0,1) + $status.Substring(1).ToLowerInvariant()
    $confidenceDisplay = $confidence

    return [PSCustomObject]@{
        ToolName = $ToolName
        ToolVersion = $ToolVersion
        RunId = $script:RunId
        CollectedAt = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
        ComputerName = $MachineInfo.ComputerName
        MachineType = $MachineInfo.MachineType
        ExecutedBy = $MachineInfo.ExecutedBy
        ExpectedUser = $MachineInfo.ExpectedUser
        ExpectedUserMatch = $MachineInfo.ExpectedUserMatch
        StageCode = $StagePriority
        StageLabel = Get-StageLabel -Code $StagePriority
        SymptomCode = $SymptomCode
        SymptomLabel = Get-SymptomLabel -Code $SymptomCode -FreeText $SymptomText
        SymptomText = $SymptomText
        ObservationMinutes = $ObservationMinutes
        SampleIntervalSeconds = $SampleIntervalSeconds
        Status = $statusDisplay
        StatusCode = $status
        Confidence = $confidenceDisplay
        ProbableCause = $probableCause
        ImpactScope = $impactScope
        SummaryPhrase = $summaryPhrase
        ProbablePerception = $probablePerception
        RecommendedAction = $recommendedAction
        ComparisonUsed = $false
        ReferenceMachine = ''
        ReferenceNote = 'Esta versão não executa comparação com referência.'
        DefenderEvaluated = $false
        DefenderNote = 'Esta versão não executa avaliação Defender/minifilter.'
        WhatDidNotIndicateFailure = $discarded.ToArray()
        Findings = $findings.ToArray()
        Signals = $signals.ToArray()
        Limitations = $limitations.ToArray()
        PassiveBenchmarkSummary = $PassiveBenchmark.SummaryPhrase
        Metrics = [PSCustomObject]@{
            SampleCount = $Timeline.SampleCount
            AverageCpuPercent = $avgCpu
            PeakCpuPercent = $peakCpu
            PeakLockFileCount = $peakLocks
            DatabaseUnavailableSamples = $dbUnavailableSamples
            NetDirUnavailableSamples = $netUnavailableSamples
            ExeUnavailableSamples = $exeUnavailableSamples
            SmbTimeoutSamples = $smbTimeoutSamples
            SeverityScore = $severityScore
            PressureLabel = $PassiveBenchmark.PressureLabel
            PeakSmbConnectionCount = $PassiveBenchmark.PeakSmbConnectionCount
            PeakSmbSessionCount = $PassiveBenchmark.PeakSmbSessionCount
            PeakRelevantOpenFileCount = $PassiveBenchmark.PeakRelevantOpenFileCount
            PeakRelevantNetDirOpenFileCount = $PassiveBenchmark.PeakRelevantNetDirOpenFileCount
        }
    }
}

function Build-SummaryText {
    param($Analysis)

    $discardLines = @()
    foreach ($line in @($Analysis.WhatDidNotIndicateFailure)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $discardLines += ('- ' + [string]$line)
        }
    }

    if ($discardLines.Count -eq 0) {
        $discardLines = @('- Ainda sem descarte confiável nesta rodada.')
    }

@"
Data/hora da coleta: $($Analysis.CollectedAt)
Máquina analisada: $($Analysis.ComputerName) — $($Analysis.MachineType)
Etapa priorizada: $($Analysis.StageLabel)
Sintoma informado: $($Analysis.SymptomLabel)
Janela de observação: $($Analysis.ObservationMinutes) minuto(s)

Status: $($Analysis.Status)
Causa mais provável: $($Analysis.ProbableCause)
Alcance do impacto: $($Analysis.ImpactScope)

Resumo:
$($Analysis.SummaryPhrase)

O que não indicou falha relevante:
$($discardLines -join [Environment]::NewLine)

Próxima ação recomendada:
$($Analysis.RecommendedAction)
"@
}

function Get-ChartDefinition {
    param($Timeline)

    $samples = @($Timeline.Samples)
    $labels = @()
    $series = @()

    foreach ($sample in $samples) {
        $labels += [string]$sample.Timestamp
    }

    $cpuData = @()
    foreach ($sample in $samples) {
        if ($null -ne $sample.CpuPercent) {
            $cpuData += [double]$sample.CpuPercent
        }
        else {
            $cpuData += $null
        }
    }
    $series += [PSCustomObject]@{
        Name = 'CPU %'
        Values = $cpuData
        Max = 100
        Color = '#2563eb'
    }

    $candidateSeries = @(
        @{ Name = 'Locks'; Property = 'LockFileCount'; Color = '#dc2626' },
        @{ Name = 'Conexões SMB'; Property = 'SmbConnectionCount'; Color = '#7c3aed' },
        @{ Name = 'Sessões SMB'; Property = 'SmbSessionCount'; Color = '#059669' },
        @{ Name = 'Open files DB'; Property = 'RelevantOpenFileCount'; Color = '#d97706' },
        @{ Name = 'Open files NetDir'; Property = 'RelevantNetDirOpenFileCount'; Color = '#ea580c' }
    )

    foreach ($candidate in $candidateSeries) {
        $values = @()
        $hasData = $false
        $maxValue = 0
        foreach ($sample in $samples) {
            $value = $sample.($candidate.Property)
            if ($null -ne $value -and "$value" -ne '') {
                try {
                    $number = [double]$value
                    $values += $number
                    if ($number -gt $maxValue) {
                        $maxValue = $number
                    }
                    if ($number -gt 0) {
                        $hasData = $true
                    }
                }
                catch {
                    $values += $null
                }
            }
            else {
                $values += $null
            }
        }

        if ($hasData -or ($candidate.Property -eq 'LockFileCount')) {
            $series += [PSCustomObject]@{
                Name = $candidate.Name
                Values = $values
                Max = $maxValue
                Color = $candidate.Color
            }
        }
    }

    $dbDropValues = @()
    $hasDbDrop = $false
    $netDropValues = @()
    $hasNetDrop = $false
    foreach ($sample in $samples) {
        $dbDrop = if ($sample.DatabaseAccessible -eq $false) { 1 } else { 0 }
        $netDrop = if ($sample.NetDirAccessible -eq $false) { 1 } else { 0 }
        $dbDropValues += $dbDrop
        $netDropValues += $netDrop
        if ($dbDrop -eq 1) { $hasDbDrop = $true }
        if ($netDrop -eq 1) { $hasNetDrop = $true }
    }

    if ($hasDbDrop) {
        $series += [PSCustomObject]@{
            Name = 'DB indisponível'
            Values = $dbDropValues
            Max = 1
            Color = '#be123c'
        }
    }

    if ($hasNetDrop) {
        $series += [PSCustomObject]@{
            Name = 'NetDir indisponível'
            Values = $netDropValues
            Max = 1
            Color = '#0f766e'
        }
    }

    return [PSCustomObject]@{
        Labels = @($labels)
        Series = @($series)
    }
}

function Build-HtmlReport {
    param(
        $Analysis,
        $MachineInfo,
        $Paths,
        $Timeline,
        $PassiveBenchmark
    )

    $statusClass = 'inconclusivo'
    switch ($Analysis.StatusCode) {
        'NORMAL' { $statusClass = 'normal' }
        'LENTO' { $statusClass = 'lento' }
        'CRÍTICO' { $statusClass = 'critico' }
        default { $statusClass = 'inconclusivo' }
    }

    $findingItems = ''
    foreach ($line in @($Analysis.Findings)) {
        $findingItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }

    $discardItems = ''
    foreach ($line in @($Analysis.WhatDidNotIndicateFailure)) {
        $discardItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }

    $limitationItems = ''
    foreach ($line in @($Analysis.Limitations)) {
        $limitationItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }

    $pathNoteItems = ''
    foreach ($line in @($Paths.Notes)) {
        $pathNoteItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($pathNoteItems)) {
        $pathNoteItems = '<li>Sem observação adicional de resolução de paths nesta rodada.</li>'
    }

    $timelineRows = ''
    foreach ($sample in @($Timeline.Samples)) {
        $cpuCell = if ($null -eq $sample.CpuPercent) { 'N/A' } else { [string]$sample.CpuPercent }
        $lockCell = if ($null -eq $sample.LockFileCount) { 'N/A' } else { [string]$sample.LockFileCount }
        $connCell = if ($null -eq $sample.SmbConnectionCount) { 'N/A' } else { [string]$sample.SmbConnectionCount }
        $sessCell = if ($null -eq $sample.SmbSessionCount) { 'N/A' } else { [string]$sample.SmbSessionCount }
        $openCell = if ($null -eq $sample.RelevantOpenFileCount) { 'N/A' } else { [string]$sample.RelevantOpenFileCount }
        $openNetCell = if ($null -eq $sample.RelevantNetDirOpenFileCount) { 'N/A' } else { [string]$sample.RelevantNetDirOpenFileCount }
        $timeoutCell = if ($sample.SmbQueryTimedOut -eq $true) { 'Sim' } else { 'Não' }

        $timelineRows += '<tr>' +
            '<td>' + (HtmlEncode ([string]$sample.Timestamp)) + '</td>' +
            '<td>' + (HtmlEncode $cpuCell) + '</td>' +
            '<td>' + (HtmlEncode $lockCell) + '</td>' +
            '<td>' + (HtmlEncode $connCell) + '</td>' +
            '<td>' + (HtmlEncode $sessCell) + '</td>' +
            '<td>' + (HtmlEncode $openCell) + '</td>' +
            '<td>' + (HtmlEncode $openNetCell) + '</td>' +
            '<td>' + (HtmlEncode ([string]$sample.DatabaseAccessible)) + '</td>' +
            '<td>' + (HtmlEncode ([string]$sample.NetDirAccessible)) + '</td>' +
            '<td>' + (HtmlEncode $timeoutCell) + '</td>' +
            '</tr>'
    }

    $chartDefinition = Get-ChartDefinition -Timeline $Timeline
    $chartJson = $chartDefinition | ConvertTo-Json -Depth 10 -Compress

    $summaryJsonPreview = [ordered]@{
        Status = $Analysis.Status
        Confidence = $Analysis.Confidence
        ProbableCause = $Analysis.ProbableCause
        SummaryPhrase = $Analysis.SummaryPhrase
        RecommendedAction = $Analysis.RecommendedAction
    }
    $summaryJsonText = $summaryJsonPreview | ConvertTo-Json -Depth 5

@"
<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ELCE ECG Diagnostics Report</title>
<style>
body { font-family: Segoe UI, Arial, sans-serif; margin: 0; background: #f3f6fb; color: #1f2937; }
.wrapper { max-width: 1220px; margin: 0 auto; padding: 24px; }
.card, details { background: #ffffff; border-radius: 14px; box-shadow: 0 2px 8px rgba(15, 23, 42, 0.08); margin-bottom: 16px; }
.card { padding: 20px; }
details { padding: 16px 20px; }
summary { cursor: pointer; font-weight: 700; }
h1, h2, h3 { margin-top: 0; }
h1 { margin-bottom: 8px; }
p { line-height: 1.5; }
.badge { display: inline-block; padding: 8px 14px; border-radius: 999px; font-weight: 700; font-size: 13px; }
.badge.normal { background: #dcfce7; color: #166534; }
.badge.lento { background: #fef3c7; color: #92400e; }
.badge.critico { background: #fee2e2; color: #991b1b; }
.badge.inconclusivo { background: #e5e7eb; color: #374151; }
.grid { display: grid; gap: 12px; grid-template-columns: repeat(2, minmax(0, 1fr)); }
.kv { border: 1px solid #e5e7eb; border-radius: 12px; padding: 10px 12px; background: #fafbfd; }
.kv strong { display: block; font-size: 12px; color: #6b7280; margin-bottom: 4px; }
.note { border-left: 4px solid #2563eb; padding: 10px 12px; background: #eff6ff; border-radius: 8px; }
.warn { border-left-color: #d97706; background: #fffbeb; }
.muted { color: #6b7280; font-size: 12px; }
ul { margin-top: 8px; }
table { width: 100%; border-collapse: collapse; font-size: 13px; }
th, td { border: 1px solid #e5e7eb; padding: 8px 10px; text-align: left; vertical-align: top; }
th { background: #f9fafb; }
.section-title { display: flex; align-items: center; justify-content: space-between; gap: 12px; }
canvas { width: 100%; height: 320px; border: 1px solid #e5e7eb; border-radius: 10px; background: #ffffff; }
.legend { display: flex; flex-wrap: wrap; gap: 12px; margin-top: 12px; }
.legend-item { display: flex; align-items: center; gap: 8px; font-size: 12px; }
.legend-color { width: 12px; height: 12px; border-radius: 999px; display: inline-block; }
pre { white-space: pre-wrap; word-break: break-word; background: #0f172a; color: #e2e8f0; padding: 14px; border-radius: 10px; overflow-x: auto; }
@media print {
  body { background: #ffffff; }
  .wrapper { max-width: none; padding: 0; }
  .card, details { box-shadow: none; border: 1px solid #d1d5db; }
}
@media (max-width: 820px) {
  .grid { grid-template-columns: 1fr; }
}
</style>
</head>
<body>
<div class="wrapper">
    <div class="card">
        <div class="section-title">
            <div>
                <span class="badge $statusClass">$([string](HtmlEncode $Analysis.Status))</span>
                <h1>$([string](HtmlEncode $ToolName))</h1>
                <p>$([string](HtmlEncode $Analysis.SummaryPhrase))</p>
            </div>
            <div class="muted">Versão $([string](HtmlEncode $ToolVersion))</div>
        </div>
        <div class="grid">
            <div class="kv"><strong>Máquina analisada</strong>$([string](HtmlEncode $Analysis.ComputerName)) — $([string](HtmlEncode $Analysis.MachineType))</div>
            <div class="kv"><strong>Data/hora da coleta</strong>$([string](HtmlEncode $Analysis.CollectedAt))</div>
            <div class="kv"><strong>Etapa priorizada</strong>$([string](HtmlEncode $Analysis.StageLabel))</div>
            <div class="kv"><strong>Sintoma informado</strong>$([string](HtmlEncode $Analysis.SymptomLabel))</div>
            <div class="kv"><strong>Causa mais provável</strong>$([string](HtmlEncode $Analysis.ProbableCause))</div>
            <div class="kv"><strong>Confiança</strong>$([string](HtmlEncode $Analysis.Confidence))</div>
            <div class="kv"><strong>Tempo de observação</strong>$([string](HtmlEncode ([string]$Analysis.ObservationMinutes))) minuto(s)</div>
            <div class="kv"><strong>Alcance do problema</strong>$([string](HtmlEncode $Analysis.ImpactScope))</div>
            <div class="kv"><strong>Comparação com referência</strong>Não — fora do escopo desta versão</div>
            <div class="kv"><strong>Defender / minifilter</strong>Não avaliado nesta versão</div>
            <div class="kv"><strong>RunId</strong>$([string](HtmlEncode $Analysis.RunId))</div>
            <div class="kv"><strong>Executado por</strong>$([string](HtmlEncode $Analysis.ExecutedBy))</div>
        </div>
    </div>

    <div class="card">
        <h2>Leitura operacional</h2>
        <p><strong>O que o usuário provavelmente percebeu:</strong> $([string](HtmlEncode $Analysis.ProbablePerception))</p>
        <p><strong>Próxima ação recomendada:</strong> $([string](HtmlEncode $Analysis.RecommendedAction))</p>
        <div class="note warn">
            <strong>Recorte desta versão:</strong> este laudo foi gerado sem comparação com referência e sem avaliação Defender/minifilter. A conclusão usa apenas contexto local, paths, índice observacional passivo e timeline temporal.
        </div>
    </div>

    <div class="card">
        <h2>O que não indicou falha relevante</h2>
        <ul>
            $discardItems
        </ul>
    </div>

    <details open>
        <summary>Seção técnica</summary>
        <div style="margin-top:16px;">
            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Achados principais</h3>
                <ul>
                    $findingItems
                </ul>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Índice observacional da rodada</h3>
                <div class="grid">
                    <div class="kv"><strong>Modo</strong>$([string](HtmlEncode $PassiveBenchmark.Mode))</div>
                    <div class="kv"><strong>Pressão da rodada</strong>$([string](HtmlEncode $PassiveBenchmark.PressureLabel))</div>
                    <div class="kv"><strong>Score de severidade</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.SeverityScore))) / 100</div>
                    <div class="kv"><strong>Amostras coletadas</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.SampleCount)))</div>
                    <div class="kv"><strong>CPU média</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.AverageCpuPercent)))</div>
                    <div class="kv"><strong>CPU pico</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.PeakCpuPercent)))</div>
                    <div class="kv"><strong>Pico de locks</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.PeakLockFileCount)))</div>
                    <div class="kv"><strong>Banco indisponível (amostras)</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.DatabaseUnavailableSamples)))</div>
                    <div class="kv"><strong>NetDir indisponível (amostras)</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.NetDirUnavailableSamples)))</div>
                    <div class="kv"><strong>Timeout SMB (amostras)</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.SmbTimeoutSamples)))</div>
                </div>
                <p class="muted">$([string](HtmlEncode $PassiveBenchmark.SummaryPhrase))</p>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Gráfico temporal — CPU + contadores relevantes</h3>
                <canvas id="timelineChart" width="1100" height="320"></canvas>
                <div id="timelineLegend" class="legend"></div>
                <p class="muted">As séries são normalizadas por escala própria para caberem no mesmo gráfico. O valor máximo real de cada série aparece na legenda.</p>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Contexto e paths efetivos</h3>
                <table>
                    <tr><th>Campo</th><th>Valor</th></tr>
                    <tr><td>Tipo da máquina</td><td>$([string](HtmlEncode $MachineInfo.MachineType))</td></tr>
                    <tr><td>Usuário esperado</td><td>$([string](HtmlEncode $MachineInfo.ExpectedUser))</td></tr>
                    <tr><td>Usuário esperado confere</td><td>$([string](HtmlEncode ([string]$MachineInfo.ExpectedUserMatch)))</td></tr>
                    <tr><td>Executável oficial</td><td>$([string](HtmlEncode $Paths.ExePath))</td></tr>
                    <tr><td>Executável acessível</td><td>$([string](HtmlEncode ([string]$Paths.ExeAccessible)))</td></tr>
                    <tr><td>Banco efetivo</td><td>$([string](HtmlEncode $Paths.DatabasePath))</td></tr>
                    <tr><td>Origem do banco</td><td>$([string](HtmlEncode $Paths.DatabasePathSource))</td></tr>
                    <tr><td>Banco acessível</td><td>$([string](HtmlEncode ([string]$Paths.DatabaseAccessible)))</td></tr>
                    <tr><td>NetDir efetivo</td><td>$([string](HtmlEncode $Paths.NetDirPath))</td></tr>
                    <tr><td>Origem do NetDir</td><td>$([string](HtmlEncode $Paths.NetDirSource))</td></tr>
                    <tr><td>NetDir acessível</td><td>$([string](HtmlEncode ([string]$Paths.NetDirAccessible)))</td></tr>
                </table>
                <h4>Observações de resolução de paths</h4>
                <ul>
                    $pathNoteItems
                </ul>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Limitações declaradas</h3>
                <ul>
                    $limitationItems
                </ul>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Timeline detalhada</h3>
                <table>
                    <tr>
                        <th>Hora</th>
                        <th>CPU %</th>
                        <th>Locks</th>
                        <th>Conexões SMB</th>
                        <th>Sessões SMB</th>
                        <th>Open DB</th>
                        <th>Open NetDir</th>
                        <th>DB OK</th>
                        <th>NetDir OK</th>
                        <th>Timeout SMB</th>
                    </tr>
                    $timelineRows
                </table>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Resumo estruturado do laudo</h3>
                <pre>$([string](HtmlEncode $summaryJsonText))</pre>
            </div>
        </div>
    </details>
</div>
<script>
(function () {
    var chartData = $chartJson;
    if (!chartData || !chartData.series || chartData.series.length === 0) {
        return;
    }

    var canvas = document.getElementById('timelineChart');
    var legend = document.getElementById('timelineLegend');
    if (!canvas || !canvas.getContext) {
        return;
    }

    var ctx = canvas.getContext('2d');
    var width = canvas.width;
    var height = canvas.height;
    var padding = { left: 56, right: 16, top: 16, bottom: 42 };
    var plotWidth = width - padding.left - padding.right;
    var plotHeight = height - padding.top - padding.bottom;
    var points = chartData.labels.length;

    function xPos(index) {
        if (points <= 1) {
            return padding.left;
        }
        return padding.left + (plotWidth * index / (points - 1));
    }

    function yPos(normalizedValue) {
        return padding.top + plotHeight - (plotHeight * normalizedValue);
    }

    ctx.clearRect(0, 0, width, height);
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, width, height);
    ctx.strokeStyle = '#e5e7eb';
    ctx.lineWidth = 1;

    for (var i = 0; i <= 4; i++) {
        var y = padding.top + (plotHeight * i / 4);
        ctx.beginPath();
        ctx.moveTo(padding.left, y);
        ctx.lineTo(width - padding.right, y);
        ctx.stroke();
    }

    ctx.beginPath();
    ctx.moveTo(padding.left, padding.top);
    ctx.lineTo(padding.left, height - padding.bottom);
    ctx.lineTo(width - padding.right, height - padding.bottom);
    ctx.strokeStyle = '#94a3b8';
    ctx.stroke();

    ctx.fillStyle = '#475569';
    ctx.font = '11px Segoe UI';
    ctx.fillText('Normalizado por escala própria', padding.left, 12);
    ctx.fillText('0', 24, height - padding.bottom + 4);
    ctx.fillText('1', 24, padding.top + 4);

    chartData.series.forEach(function (series) {
        var localMax = Number(series.max || 0);
        if (!localMax || localMax < 1) {
            if (series.name === 'CPU %') {
                localMax = 100;
            } else {
                localMax = 1;
            }
        }

        ctx.beginPath();
        var started = false;
        for (var idx = 0; idx < series.values.length; idx++) {
            var rawValue = series.values[idx];
            if (rawValue === null || rawValue === undefined || rawValue === '') {
                continue;
            }
            var numberValue = Number(rawValue);
            var normalized = numberValue / localMax;
            if (normalized < 0) { normalized = 0; }
            if (normalized > 1) { normalized = 1; }
            var x = xPos(idx);
            var y = yPos(normalized);
            if (!started) {
                ctx.moveTo(x, y);
                started = true;
            } else {
                ctx.lineTo(x, y);
            }
        }
        ctx.strokeStyle = series.color;
        ctx.lineWidth = 2;
        ctx.stroke();
    });

    chartData.labels.forEach(function (label, index) {
        if ((index % 3) === 0 || index === chartData.labels.length - 1) {
            var x = xPos(index);
            ctx.save();
            ctx.translate(x, height - padding.bottom + 14);
            ctx.rotate(-0.45);
            ctx.fillStyle = '#475569';
            ctx.font = '10px Segoe UI';
            ctx.fillText(label, 0, 0);
            ctx.restore();
        }
    });

    chartData.series.forEach(function (series) {
        var localMax = Number(series.max || 0);
        var legendItem = document.createElement('div');
        legendItem.className = 'legend-item';
        var colorBox = document.createElement('span');
        colorBox.className = 'legend-color';
        colorBox.style.backgroundColor = series.color;
        var text = document.createElement('span');
        text.textContent = series.name + ' (máx=' + localMax + ')';
        legendItem.appendChild(colorBox);
        legendItem.appendChild(text);
        legend.appendChild(legendItem);
    });
})();
</script>
</body>
</html>
"@
}

function Copy-LatestArtifacts {
    param(
        [string]$RunRoot,
        [string]$LatestRoot
    )

    Log 'Atualizando Latest.' 'STEP'

    Copy-Item -Path (Join-Path $RunRoot 'ELCE_ECG_Diagnostics_Report.html') -Destination (Join-Path $LatestRoot 'ELCE_ECG_Diagnostics_Report.html') -Force
    if (Test-Path (Join-Path $RunRoot 'ELCE_ECG_Diagnostics_Summary.txt')) {
        Copy-Item -Path (Join-Path $RunRoot 'ELCE_ECG_Diagnostics_Summary.txt') -Destination (Join-Path $LatestRoot 'ELCE_ECG_Diagnostics_Summary.txt') -Force
    }
    if (Test-Path (Join-Path $RunRoot 'ELCE_ECG_Diagnostics_Summary.json')) {
        Copy-Item -Path (Join-Path $RunRoot 'ELCE_ECG_Diagnostics_Summary.json') -Destination (Join-Path $LatestRoot 'ELCE_ECG_Diagnostics_Summary.json') -Force
    }
    Write-Utf8NoBomFile -Path (Join-Path $LatestRoot 'latest_run.txt') -Content $script:RunId
}

function Load-ParameterFileIfPresent {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return
    }
    if (-not (Test-Path $Path)) {
        return
    }

    try {
        $loaded = ConvertFrom-Json -InputObject (Get-Content $Path -Raw)
        if ($null -ne $loaded) {
            if ($loaded.PSObject.Properties.Name -contains 'StagePriority') {
                $candidate = [string]$loaded.StagePriority
                if (@('ABRIR_EXAME','SALVAR_FINALIZAR','GERAL') -contains $candidate) {
                    $script:StagePriority = $candidate
                }
            }
            if ($loaded.PSObject.Properties.Name -contains 'SymptomCode') {
                $candidateSymptom = [string]$loaded.SymptomCode
                if (@('LENTIDAO_TRAVAMENTO','INCONCLUSIVO') -contains $candidateSymptom) {
                    $script:SymptomCode = $candidateSymptom
                }
            }
            if ($loaded.PSObject.Properties.Name -contains 'SymptomText' -and $null -ne $loaded.SymptomText) {
                $script:SymptomText = [string]$loaded.SymptomText
            }
            if ($loaded.PSObject.Properties.Name -contains 'ObservationMinutes') {
                try {
                    $minutes = [int]$loaded.ObservationMinutes
                    if ($minutes -ge 1 -and $minutes -le 120) {
                        $script:ObservationMinutes = $minutes
                    }
                }
                catch {}
            }
            if ($loaded.PSObject.Properties.Name -contains 'SampleIntervalSeconds') {
                try {
                    $seconds = [int]$loaded.SampleIntervalSeconds
                    if ($seconds -ge 5 -and $seconds -le 300) {
                        $script:SampleIntervalSeconds = $seconds
                    }
                }
                catch {}
            }
            if ($loaded.PSObject.Properties.Name -contains 'OpenReportOnSuccess' -and $loaded.OpenReportOnSuccess -eq $true) {
                $script:OpenReportOnSuccess = $true
            }
        }
    }
    catch {
        Write-Host ('Falha ao carregar ParameterFile: ' + $_.Exception.Message)
    }
}

foreach ($dir in @($OutputRoot, $RunsRoot, $LatestRoot)) {
    if (-not (Test-Path $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
}

Load-ParameterFileIfPresent -Path $ParameterFile

$script:RunId = (Get-Date -Format 'dd-MM-yyyy_HH-mm-ss') + '_' + $env:COMPUTERNAME
$script:RunRoot = Join-Path $RunsRoot $script:RunId
if (-not (Test-Path $script:RunRoot)) {
    New-Item -Path $script:RunRoot -ItemType Directory -Force | Out-Null
}
$script:LogFile = Join-Path $script:RunRoot 'execution.log'
Write-Utf8NoBomFile -Path $script:LogFile -Content ''

try {
    Log ($ToolName + ' iniciado. Versão=' + $ToolVersion) 'START'
    Log ('RunId=' + $script:RunId) 'INFO'
    Log ('Etapa=' + (Get-StageLabel -Code $StagePriority) + ' | Sintoma=' + (Get-SymptomLabel -Code $SymptomCode -FreeText $SymptomText)) 'INFO'
    Log ('Janela=' + $ObservationMinutes + ' min | Intervalo=' + $SampleIntervalSeconds + ' s') 'INFO'

    Log 'Resolvendo contexto da máquina.' 'STEP'
    $machineInfo = Get-KnownMachineInfo -ComputerName $env:COMPUTERNAME

    Log 'Resolvendo paths oficiais e efetivos.' 'STEP'
    $paths = Resolve-EcgPaths

    $context = [PSCustomObject]@{
        ToolName = $ToolName
        ToolVersion = $ToolVersion
        RunId = $script:RunId
        CollectedAt = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
        Machine = $machineInfo
        Paths = $paths
        Input = [PSCustomObject]@{
            StageCode = $StagePriority
            StageLabel = Get-StageLabel -Code $StagePriority
            SymptomCode = $SymptomCode
            SymptomLabel = Get-SymptomLabel -Code $SymptomCode -FreeText $SymptomText
            SymptomText = $SymptomText
            ObservationMinutes = $ObservationMinutes
            SampleIntervalSeconds = $SampleIntervalSeconds
            OpenReportOnSuccess = [bool]$OpenReportOnSuccess
        }
    }

    Log 'Gravando context.json.' 'STEP'
    Save-JsonFile -Path (Join-Path $script:RunRoot 'context.json') -Object $context

    Log 'Iniciando coleta temporal.' 'STEP'
    $timeline = Collect-Timeline -MachineInfo $machineInfo -Paths $paths -Minutes $ObservationMinutes -IntervalSeconds $SampleIntervalSeconds
    Log ('Coleta temporal concluída. Amostras=' + [string]$timeline.SampleCount) 'STEP'

    Log 'Gravando timeline.json.' 'STEP'
    Save-JsonFile -Path (Join-Path $script:RunRoot 'timeline.json') -Object $timeline

    Log 'Calculando índice observacional da rodada.' 'STEP'
    $passiveBenchmark = Build-PassiveBenchmark -Timeline $timeline -StagePriority $StagePriority -SymptomCode $SymptomCode
    Log ('Índice observacional calculado. Score=' + [string]$passiveBenchmark.SeverityScore) 'STEP'

    Log 'Gravando benchmark.json.' 'STEP'
    Save-JsonFile -Path (Join-Path $script:RunRoot 'benchmark.json') -Object $passiveBenchmark

    Log 'Iniciando análise.' 'STEP'
    $analysis = Build-AnalysisModel -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $passiveBenchmark
    Log ('Análise concluída. Status=' + [string]$analysis.Status + ' | Causa=' + [string]$analysis.ProbableCause) 'STEP'

    Log 'Gravando analysis.json.' 'STEP'
    Save-JsonFile -Path (Join-Path $script:RunRoot 'analysis.json') -Object $analysis

    Log 'Montando HTML principal.' 'STEP'
    $reportHtml = Build-HtmlReport -Analysis $analysis -MachineInfo $machineInfo -Paths $paths -Timeline $timeline -PassiveBenchmark $passiveBenchmark

    Log 'Gravando ELCE_ECG_Diagnostics_Report.html.' 'STEP'
    Write-Utf8NoBomFile -Path (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Report.html') -Content ([string]$reportHtml)

    try {
        Log 'Montando Summary secundário.' 'STEP'
        $summaryText = Build-SummaryText -Analysis $analysis
        Write-Utf8NoBomFile -Path (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Summary.txt') -Content ([string]$summaryText)
        Save-JsonFile -Path (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Summary.json') -Object $analysis
    }
    catch {
        Log ('Summary secundário falhou sem abortar a rodada: ' + $_.Exception.Message) 'WARN'
    }

    Copy-LatestArtifacts -RunRoot $script:RunRoot -LatestRoot $LatestRoot

    Log ('Laudo salvo em: ' + (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Report.html')) 'INFO'
    Log 'Rodada finalizada com sucesso.' 'END'

    if ($OpenReportOnSuccess -and (Test-Path (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Report.html'))) {
        try {
            Start-Process -FilePath (Join-Path $script:RunRoot 'ELCE_ECG_Diagnostics_Report.html') | Out-Null
        }
        catch {
            Log ('Falha ao abrir o HTML automaticamente: ' + $_.Exception.Message) 'WARN'
        }
    }

    exit 0
}
catch {
    Log ('Falha fatal: ' + $_.Exception.Message) 'ERROR'
    if ($null -ne $_.InvocationInfo) {
        Log ('Falha fatal em: ' + [string]$_.InvocationInfo.PositionMessage) 'ERROR'
    }
    if ($null -ne $_.ScriptStackTrace) {
        Log ('Stack: ' + [string]$_.ScriptStackTrace) 'ERROR'
    }
    try {
        $fallback = [PSCustomObject]@{
            ToolName = $ToolName
            ToolVersion = $ToolVersion
            RunId = $script:RunId
            CollectedAt = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
            Status = 'Inconclusivo'
            StatusCode = 'INCONCLUSIVO'
            Confidence = 'Baixa'
            ProbableCause = 'Ainda sem causa definida'
            SummaryPhrase = 'A rodada falhou antes de concluir o laudo técnico.'
            RecommendedAction = 'Revisar execution.log desta rodada e rerodar a ferramenta.'
            Error = $_.Exception.Message
        }
        Save-JsonFile -Path (Join-Path $script:RunRoot 'analysis.json') -Object $fallback
    }
    catch {}
    exit 1
}
