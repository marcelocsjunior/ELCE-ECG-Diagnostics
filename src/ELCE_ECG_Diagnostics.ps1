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


$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$UnitProfilesFile = Join-Path $ScriptRoot 'ECG_UnitProfiles.json'

function Restart-ScriptWithBypassIfNeeded {
    if (-not $MyInvocation.MyCommand.Path) { return }
    if ($env:ELCE_ECG_BYPASS_RESTARTED -eq '1') { return }

    $effectivePolicy = $null
    try { $effectivePolicy = Get-ExecutionPolicy -ErrorAction Stop } catch { $effectivePolicy = $null }

    if ($effectivePolicy -in @('Restricted','AllSigned')) {
        Write-Host "Política de execução '$effectivePolicy' detectada. Reiniciando com ExecutionPolicy Bypass..." -ForegroundColor Yellow
        $args = New-Object System.Collections.Generic.List[string]
        $args.Add('-NoProfile')
        $args.Add('-ExecutionPolicy')
        $args.Add('Bypass')
        $args.Add('-File')
        $args.Add($MyInvocation.MyCommand.Path)

        if ($PSBoundParameters.ContainsKey('StagePriority')) { $args.Add('-StagePriority'); $args.Add($StagePriority) }
        if ($PSBoundParameters.ContainsKey('SymptomCode')) { $args.Add('-SymptomCode'); $args.Add($SymptomCode) }
        if (-not [string]::IsNullOrWhiteSpace($SymptomText)) { $args.Add('-SymptomText'); $args.Add($SymptomText) }
        if ($PSBoundParameters.ContainsKey('ObservationMinutes')) { $args.Add('-ObservationMinutes'); $args.Add([string]$ObservationMinutes) }
        if ($PSBoundParameters.ContainsKey('SampleIntervalSeconds')) { $args.Add('-SampleIntervalSeconds'); $args.Add([string]$SampleIntervalSeconds) }
        if (-not [string]::IsNullOrWhiteSpace($ParameterFile)) { $args.Add('-ParameterFile'); $args.Add($ParameterFile) }
        if ($OpenReportOnSuccess) { $args.Add('-OpenReportOnSuccess') }

        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.Arguments = ($args | ForEach-Object {
            if ($_ -match '\s') { '"' + ($_ -replace '"', '""') + '"' } else { $_ }
        }) -join ' '
        $psi.UseShellExecute = $false
        $psi.EnvironmentVariables['ELCE_ECG_BYPASS_RESTARTED'] = '1'

        $process = [System.Diagnostics.Process]::Start($psi)
        $process.WaitForExit()
        exit $process.ExitCode
    }
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
      "computerPatterns": ["^ELCUN1-"],
      "examStationPatterns": ["^ELCUN1-ECG"],
      "viewerStationPatterns": ["^ELCUN1-(CST|CON|VIEW)"],
      "expectedUserByHost": {
        "ELCUN1-ECG": "elce\\ecg.un1",
        "ELCUN1-CST2": "elce\\ewaldo.bayao"
      }
    },
    "UN2": {
      "name": "Unidade 2",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN2-ECG\\hw",
      "netDirPath": "\\\\ELCUN2-ECG\\hw\\NetDir",
      "fallbackDbPath": "",
      "fileServerHost": "",
      "computerPatterns": ["^ELCUN2-"],
      "examStationPatterns": ["^ELCUN2-ECG"],
      "viewerStationPatterns": ["^ELCUN2-(CST|CON|VIEW)"],
      "expectedUserByHost": {}
    },
    "UN3": {
      "name": "Unidade 3",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN3-ECG\\hw",
      "netDirPath": "\\\\ELCUN3-ECG\\hw\\NetDir",
      "fallbackDbPath": "",
      "fileServerHost": "",
      "computerPatterns": ["^ELCUN3-"],
      "examStationPatterns": ["^ELCUN3-ECG"],
      "viewerStationPatterns": ["^ELCUN3-(CST|CON|VIEW)"],
      "expectedUserByHost": {}
    }
  }
}
'@
    return ($json | ConvertFrom-Json)
}

function Get-UnitProfiles {
    $defaults = Get-DefaultUnitProfiles
    if (Test-Path -LiteralPath $UnitProfilesFile) {
        try {
            $external = Get-Content -LiteralPath $UnitProfilesFile -Raw | ConvertFrom-Json
            if ($null -ne $external -and $null -ne $external.units) {
                return $external
            }
        }
        catch {}
    }
    return $defaults
}

function Get-NormalizedPathForCompare {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) { return '' }

    $tmp = ($Value -replace '/', '\').Trim()
    while ($tmp.Length -gt 3 -and $tmp.EndsWith('\')) {
        $tmp = $tmp.Substring(0, $tmp.Length - 1)
    }
    return $tmp.ToLowerInvariant()
}

function Test-PatternMatch {
    param(
        [string]$Value,
        [object]$Patterns
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or $null -eq $Patterns) { return $false }

    foreach ($pattern in @($Patterns)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$pattern) -and $Value -match [string]$pattern) {
            return $true
        }
    }

    return $false
}

function Get-DetectedUnitCode {
    param(
        [string]$ComputerName,
        $ProfileConfig
    )

    $computerUpper = ([string]$ComputerName).ToUpperInvariant()

    if ($null -ne $ProfileConfig -and $null -ne $ProfileConfig.units) {
        foreach ($property in $ProfileConfig.units.PSObject.Properties) {
            $unitCode = [string]$property.Name
            $profile = $property.Value
            if ($null -ne $profile -and (Test-PatternMatch -Value $computerUpper -Patterns $profile.computerPatterns)) {
                return $unitCode
            }
        }
    }

    if ($computerUpper -match '^ELCUN(\d+)-') {
        return ('UN' + $matches[1])
    }

    return 'UNKNOWN'
}

function Get-UnitProfile {
    param(
        [string]$UnitCode,
        $ProfileConfig
    )

    if ([string]::IsNullOrWhiteSpace($UnitCode) -or $null -eq $ProfileConfig -or $null -eq $ProfileConfig.units) {
        return $null
    }

    if ($ProfileConfig.units.PSObject.Properties.Name -contains $UnitCode) {
        return $ProfileConfig.units.$UnitCode
    }

    return $null
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

Restart-ScriptWithBypassIfNeeded
$script:UnitProfiles = Get-UnitProfiles

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ToolName = 'ELCE ECG Diagnostics'
$ToolVersion = '3.6-multiunit-bde-contract'
$ToolRoot = 'C:\ECG\Tool'
$OutputRoot = 'C:\ECG\Output'
$RunsRoot = Join-Path $OutputRoot 'Runs'
$LatestRoot = Join-Path $OutputRoot 'Latest'
$OfficialExePath = [string]$script:UnitProfiles.defaultExePath

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

    $computerUpper = ([string]$ComputerName).ToUpperInvariant()
    $profileConfig = $script:UnitProfiles
    $unitCode = Get-DetectedUnitCode -ComputerName $computerUpper -ProfileConfig $profileConfig
    $profile = Get-UnitProfile -UnitCode $unitCode -ProfileConfig $profileConfig
    $machineType = 'Estação Windows'
    $expectedUser = ''
    $windowsProductType = $null
    $profileName = if ($null -ne $profile -and -not [string]::IsNullOrWhiteSpace([string]$profile.name)) { [string]$profile.name } else { 'Sem perfil dedicado' }
    $topologyType = if ($null -ne $profile -and -not [string]::IsNullOrWhiteSpace([string]$profile.topology)) { [string]$profile.topology } else { 'AUTO_REGISTRY_ENV' }
    $fileServerHost = if ($null -ne $profile -and -not [string]::IsNullOrWhiteSpace([string]$profile.fileServerHost)) { [string]$profile.fileServerHost } else { '' }

    try {
        $windowsProductType = [int](Get-CimInstance Win32_OperatingSystem -ErrorAction Stop).ProductType
    }
    catch {
        try {
            $windowsProductType = [int](Get-WmiObject Win32_OperatingSystem -ErrorAction Stop).ProductType
        }
        catch {
            $windowsProductType = $null
        }
    }

    if ($null -ne $profile -and $null -ne $profile.expectedUserByHost -and $profile.expectedUserByHost.PSObject.Properties.Name -contains $computerUpper) {
        $expectedUser = [string]$profile.expectedUserByHost.$computerUpper
    }

    if ($null -ne $profile -and (Test-PatternMatch -Value $computerUpper -Patterns $profile.examStationPatterns)) {
        $machineType = 'Estação de exames'
    }
    elseif ($null -ne $profile -and (Test-PatternMatch -Value $computerUpper -Patterns $profile.viewerStationPatterns)) {
        $machineType = 'Estação de visualização'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($fileServerHost) -and $computerUpper -eq $fileServerHost.ToUpperInvariant()) {
        $machineType = 'Servidor de arquivos'
    }
    elseif ($computerUpper -match '(^|[-_])ECG([0-9A-Z_-]*$)') {
        $machineType = 'Estação de exames'
    }
    elseif ($computerUpper -match '(^|[-_])(CST|CON|VIEW)([0-9A-Z_-]*$)') {
        $machineType = 'Estação de visualização'
    }
    elseif (($windowsProductType -in @(2,3)) -and ($computerUpper -match 'FS|FILE|SRV')) {
        $machineType = 'Servidor de arquivos'
    }
    elseif ($computerUpper -match '(^|[-_])(VMW|VM|VDI|WKS|WS|PC)([0-9A-Z_-]*$)') {
        $machineType = 'Estação de trabalho'
    }
    elseif ($windowsProductType -in @(2,3)) {
        $machineType = 'Servidor Windows'
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
        UnitCode = $unitCode
        ProfileName = $profileName
        TopologyType = $topologyType
        FileServerHost = $fileServerHost
        MachineType = $machineType
        ExecutedBy = $executedBy
        ExpectedUser = $expectedUser
        ExpectedUserMatch = $expectedUserMatch
        WindowsProductType = $windowsProductType
    }
}

function Resolve-EcgPaths {
    param($MachineInfo)

    $profile = Get-UnitProfile -UnitCode $MachineInfo.UnitCode -ProfileConfig $script:UnitProfiles
    $pathNotes = New-Object System.Collections.Generic.List[string]

    $profileDb = ''
    $profileNetDir = ''
    $profileFallbackDb = ''
    if ($null -ne $profile) {
        if (-not [string]::IsNullOrWhiteSpace([string]$profile.dbPath)) { $profileDb = [string]$profile.dbPath }
        if (-not [string]::IsNullOrWhiteSpace([string]$profile.netDirPath)) { $profileNetDir = [string]$profile.netDirPath }
        if (-not [string]::IsNullOrWhiteSpace([string]$profile.fallbackDbPath)) { $profileFallbackDb = [string]$profile.fallbackDbPath }
    }

    $envProcessHwDb = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::Process)
    $envUserHwDb = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::User)
    $envMachineHwDb = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::Machine)

    $regDb = $null
    foreach ($candidate in @(
        'HKCU:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\Software\WOW6432Node\HeartWare\ECGV6\Geral'
    )) {
        $tmp = Get-RegistryValueSafe -Path $candidate -Name 'Caminho Database'
        if (-not [string]::IsNullOrWhiteSpace([string]$tmp)) {
            $regDb = [string]$tmp
            break
        }
    }

    $candidateDbList = @()
    if (-not [string]::IsNullOrWhiteSpace($profileDb)) { $candidateDbList += [PSCustomObject]@{ Path = $profileDb; Source = 'Perfil da unidade' } }
    if (-not [string]::IsNullOrWhiteSpace($envProcessHwDb)) { $candidateDbList += [PSCustomObject]@{ Path = [string]$envProcessHwDb; Source = 'HW_CAMINHO_DB (Process)' } }
    if (-not [string]::IsNullOrWhiteSpace($envUserHwDb)) { $candidateDbList += [PSCustomObject]@{ Path = [string]$envUserHwDb; Source = 'HW_CAMINHO_DB (User)' } }
    if (-not [string]::IsNullOrWhiteSpace($envMachineHwDb)) { $candidateDbList += [PSCustomObject]@{ Path = [string]$envMachineHwDb; Source = 'HW_CAMINHO_DB (Machine)' } }
    if (-not [string]::IsNullOrWhiteSpace($regDb)) { $candidateDbList += [PSCustomObject]@{ Path = [string]$regDb; Source = 'Registro ECG/BDE' } }
    if (-not [string]::IsNullOrWhiteSpace($profileFallbackDb)) { $candidateDbList += [PSCustomObject]@{ Path = $profileFallbackDb; Source = 'Fallback de perfil' } }

    $effectiveDb = ''
    $dbSource = 'Não determinado'
    foreach ($candidate in $candidateDbList) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate.Path) -and (Test-Path -LiteralPath ([string]$candidate.Path))) {
            $effectiveDb = [string]$candidate.Path
            $dbSource = [string]$candidate.Source
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace($effectiveDb) -and $candidateDbList.Count -gt 0) {
        $effectiveDb = [string]$candidateDbList[0].Path
        $dbSource = ([string]$candidateDbList[0].Source) + ' (não acessível nesta rodada)'
    }

    $expectedDb = if (-not [string]::IsNullOrWhiteSpace($profileDb)) { $profileDb } else { $effectiveDb }
    $expectedNetDir = ''
    if (-not [string]::IsNullOrWhiteSpace($profileNetDir)) {
        $expectedNetDir = $profileNetDir
    }
    elseif (-not [string]::IsNullOrWhiteSpace($expectedDb)) {
        $expectedNetDir = Join-Path $expectedDb 'NetDir'
    }

    $currentBdeNetDir = Get-BdeNetDirFromRegistry
    $effectiveNetDir = $expectedNetDir
    $netDirSource = 'Perfil/contrato'
    if ([string]::IsNullOrWhiteSpace($effectiveNetDir) -and -not [string]::IsNullOrWhiteSpace($currentBdeNetDir)) {
        $effectiveNetDir = $currentBdeNetDir
        $netDirSource = 'Registro BDE'
    }
    elseif (-not [string]::IsNullOrWhiteSpace($effectiveNetDir) -and -not (Test-Path -LiteralPath $effectiveNetDir) -and -not [string]::IsNullOrWhiteSpace($currentBdeNetDir) -and (Test-Path -LiteralPath $currentBdeNetDir)) {
        $effectiveNetDir = $currentBdeNetDir
        $netDirSource = 'Registro BDE (fallback de acessibilidade)'
    }

    $bdeNetDirStatus = 'NAO_DETERMINADO'
    if ([string]::IsNullOrWhiteSpace($currentBdeNetDir) -and [string]::IsNullOrWhiteSpace($expectedNetDir)) {
        $bdeNetDirStatus = 'NAO_DETERMINADO'
    }
    elseif ([string]::IsNullOrWhiteSpace($currentBdeNetDir)) {
        $bdeNetDirStatus = 'AUSENTE'
    }
    elseif ([string]::IsNullOrWhiteSpace($expectedNetDir)) {
        $bdeNetDirStatus = 'SEM_EXPECTATIVA'
    }
    elseif ((Get-NormalizedPathForCompare $currentBdeNetDir) -eq (Get-NormalizedPathForCompare $expectedNetDir)) {
        $bdeNetDirStatus = 'OK'
    }
    else {
        $bdeNetDirStatus = 'DIVERGENTE'
    }

    $lockControlFilePresent = $false
    foreach ($lockCandidate in @($effectiveNetDir, $expectedNetDir, $currentBdeNetDir)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$lockCandidate) -and (Test-Path -LiteralPath $lockCandidate)) {
            if (Test-Path -LiteralPath (Join-Path ([string]$lockCandidate) 'PDOXUSRS.NET')) {
                $lockControlFilePresent = $true
                break
            }
        }
    }

    if ([string]::IsNullOrWhiteSpace($effectiveDb)) {
        $pathNotes.Add('Nenhum caminho de banco pôde ser resolvido a partir do perfil, ambiente ou registro.')
    }
    elseif ($dbSource -match 'não acessível') {
        $pathNotes.Add('Foi resolvido um caminho de banco lógico, porém ele não estava acessível nesta rodada.')
    }

    if ($bdeNetDirStatus -eq 'AUSENTE') {
        $pathNotes.Add('NETDIR ausente no registro do BDE.')
    }
    elseif ($bdeNetDirStatus -eq 'DIVERGENTE') {
        $pathNotes.Add('NETDIR do BDE diverge do contrato esperado para a unidade.')
    }

    return [PSCustomObject]@{
        UnitCode = $MachineInfo.UnitCode
        ProfileName = $MachineInfo.ProfileName
        TopologyType = $MachineInfo.TopologyType
        FileServerHost = $MachineInfo.FileServerHost
        ExePath = $OfficialExePath
        ExeAccessible = [bool](Test-Path -LiteralPath $OfficialExePath)
        DatabasePath = $effectiveDb
        DatabasePathExpected = $expectedDb
        DatabasePathSource = $dbSource
        DatabaseAccessible = [bool](-not [string]::IsNullOrWhiteSpace($effectiveDb) -and (Test-Path -LiteralPath $effectiveDb))
        NetDirPath = $effectiveNetDir
        NetDirPathExpected = $expectedNetDir
        NetDirSource = $netDirSource
        NetDirAccessible = [bool](-not [string]::IsNullOrWhiteSpace($effectiveNetDir) -and (Test-Path -LiteralPath $effectiveNetDir))
        CurrentBdeNetDir = $currentBdeNetDir
        BdeNetDirStatus = $bdeNetDirStatus
        LockControlFilePresent = [bool]$lockControlFilePresent
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
    $deadline = $started.AddMinutes($Minutes)
    $sampleTarget = [int][math]::Ceiling(($Minutes * 60) / $IntervalSeconds)
    if ($sampleTarget -lt 1) {
        $sampleTarget = 1
    }

    $targetWallClockSeconds = [math]::Round(($Minutes * 60), 2)
    $stoppedByWallClock = $false

    Log ('Coleta temporal iniciada. Janela=' + $Minutes + ' min | Intervalo=' + $IntervalSeconds + ' s | Meta de amostras=' + $sampleTarget + ' | Modo=wall-clock bounded') 'STEP'

    for ($i = 1; $i -le $sampleTarget; $i++) {
        $iterationPlannedStart = $started.AddSeconds(($i - 1) * $IntervalSeconds)
        $iterationNow = Get-Date

        if (($i -gt 1) -and ($iterationNow -ge $deadline)) {
            $stoppedByWallClock = $true
            Log ('Watchdog temporal encerrou a coleta antes da amostra ' + $i + '/' + $sampleTarget + '. Janela configurada já foi consumida no relógio real.') 'WARN'
            break
        }

        $startedLateMs = [math]::Round([math]::Max(0, ($iterationNow - $iterationPlannedStart).TotalMilliseconds), 0)
        $sampleStartedAt = $iterationNow
        $sample = Get-TimelineSample -MachineInfo $MachineInfo -Paths $Paths -SampleIndex $i
        $sampleEndedAt = Get-Date
        $sampleDurationMs = [math]::Round(($sampleEndedAt - $sampleStartedAt).TotalMilliseconds, 0)

        try {
            $sample | Add-Member -NotePropertyName PlannedStartIso -NotePropertyValue ($iterationPlannedStart.ToString('s')) -Force
            $sample | Add-Member -NotePropertyName SampleDurationMs -NotePropertyValue $sampleDurationMs -Force
            $sample | Add-Member -NotePropertyName StartedLateMs -NotePropertyValue $startedLateMs -Force
        }
        catch {}

        $samples += $sample

        $cpuText = 'N/A'
        if ($null -ne $sample.CpuPercent) {
            $cpuText = ([string]$sample.CpuPercent) + '%'
        }

        $dbText = [string]$sample.DatabaseAccessible
        $netText = [string]$sample.NetDirAccessible
        $lockText = if ($null -eq $sample.LockFileCount) { 'N/A' } else { [string]$sample.LockFileCount }
        Log ('Amostra ' + $i + '/' + $sampleTarget + ' | Hora=' + $sample.Timestamp + ' | CPU=' + $cpuText + ' | Locks=' + $lockText + ' | DB=' + $dbText + ' | NetDir=' + $netText + ' | DuracaoMs=' + [string]$sampleDurationMs + ' | AtrasoInicioMs=' + [string]$startedLateMs) 'INFO'

        if ($i -lt $sampleTarget) {
            $nextPlannedStart = $started.AddSeconds($i * $IntervalSeconds)
            $remainingMs = [int][math]::Floor(($nextPlannedStart - (Get-Date)).TotalMilliseconds)

            if (($remainingMs -gt 0) -and ((Get-Date) -lt $deadline)) {
                Start-Sleep -Milliseconds $remainingMs
            }
            elseif ($remainingMs -le -250) {
                Log ('Amostra ' + $i + ' encerrou com overrun de ' + [string]([math]::Abs($remainingMs)) + ' ms em relacao ao cronograma planejado.') 'WARN'
            }
        }
    }

    $ended = Get-Date
    $actualWallClockSeconds = [math]::Round(($ended - $started).TotalSeconds, 2)
    $timingDriftSeconds = [math]::Round(($actualWallClockSeconds - $targetWallClockSeconds), 2)

    $timingIntegrity = 'OK'
    if ($actualWallClockSeconds -gt ($targetWallClockSeconds + ($IntervalSeconds * 2))) {
        $timingIntegrity = 'Crítica'
    }
    elseif (($actualWallClockSeconds -gt ($targetWallClockSeconds + $IntervalSeconds)) -or ($samples.Count -lt $sampleTarget)) {
        $timingIntegrity = 'Atenção'
    }

    $timingNote = 'Janela real respeitou o orçamento temporal configurado.'
    if ($timingIntegrity -eq 'Atenção') {
        $timingNote = 'Janela real excedeu levemente o orçamento temporal configurado ou exigiu encerramento antecipado por watchdog.'
    }
    elseif ($timingIntegrity -eq 'Crítica') {
        $timingNote = 'Janela real excedeu de forma relevante o orçamento temporal configurado; investigar custo por amostra e bloqueios internos.'
    }

    return [PSCustomObject]@{
        StartedAt = $started.ToString('dd/MM/yyyy HH:mm:ss')
        EndedAt = $ended.ToString('dd/MM/yyyy HH:mm:ss')
        ObservationMinutes = $Minutes
        IntervalSeconds = $IntervalSeconds
        RequestedSampleTarget = $sampleTarget
        SampleCount = $samples.Count
        Samples = @($samples)
        SchedulingMode = 'WALL_CLOCK_BOUNDED'
        StoppedByWallClock = $stoppedByWallClock
        TargetWallClockSeconds = $targetWallClockSeconds
        ActualWallClockSeconds = $actualWallClockSeconds
        ActualWallClockMinutes = [math]::Round(($actualWallClockSeconds / 60), 2)
        TimingDriftSeconds = $timingDriftSeconds
        TimingIntegrity = $timingIntegrity
        TimingNote = $timingNote
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

    $timingIntegrity = 'OK'
    $timingNote = 'Janela real respeitou o orçamento temporal configurado.'
    $targetWallClockSeconds = $null
    $actualWallClockSeconds = $null
    $actualWallClockMinutes = $null
    $timingDriftSeconds = $null
    $requestedSampleTarget = $sampleCount
    $stoppedByWallClock = $false

    if ($Timeline.PSObject.Properties.Name -contains 'TimingIntegrity') {
        $timingIntegrity = [string]$Timeline.TimingIntegrity
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TimingNote') {
        $timingNote = [string]$Timeline.TimingNote
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TargetWallClockSeconds') {
        $targetWallClockSeconds = $Timeline.TargetWallClockSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'ActualWallClockSeconds') {
        $actualWallClockSeconds = $Timeline.ActualWallClockSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'ActualWallClockMinutes') {
        $actualWallClockMinutes = $Timeline.ActualWallClockMinutes
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TimingDriftSeconds') {
        $timingDriftSeconds = $Timeline.TimingDriftSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'RequestedSampleTarget') {
        $requestedSampleTarget = $Timeline.RequestedSampleTarget
    }
    if ($Timeline.PSObject.Properties.Name -contains 'StoppedByWallClock') {
        $stoppedByWallClock = [bool]$Timeline.StoppedByWallClock
    }

    $summaryPhrase = 'Índice observacional da rodada calculado de forma passiva, sem benchmark assistido dentro do ECG.'
    if ($timingIntegrity -ne 'OK') {
        $summaryPhrase += ' Integridade temporal da coleta: ' + $timingIntegrity + '. ' + $timingNote
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
        RequestedSampleTarget = $requestedSampleTarget
        SampleCount = $Timeline.SampleCount
        SchedulingMode = $Timeline.SchedulingMode
        StoppedByWallClock = $stoppedByWallClock
        TargetWallClockSeconds = $targetWallClockSeconds
        ActualWallClockSeconds = $actualWallClockSeconds
        ActualWallClockMinutes = $actualWallClockMinutes
        TimingDriftSeconds = $timingDriftSeconds
        TimingIntegrity = $timingIntegrity
        TimingNote = $timingNote
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
        SummaryPhrase = $summaryPhrase
    }
}


function Build-AnalysisModel {
    param(
        $MachineInfo,
        $Paths,
        $Timeline,
        $PassiveBenchmark
    )

    function New-HypothesisBucket {
        param(
            [string]$Key,
            [string]$Display,
            [string]$ImpactScope,
            [string]$RecommendedAction
        )

        return [ordered]@{
            Key = $Key
            Display = $Display
            ImpactScope = $ImpactScope
            RecommendedAction = $RecommendedAction
            Score = 0
            Evidence = (New-Object System.Collections.Generic.List[string])
            CounterEvidence = (New-Object System.Collections.Generic.List[string])
        }
    }

    function Add-HypothesisEvidence {
        param(
            $Bucket,
            [int]$Points,
            [string]$Text
        )

        $Bucket.Score += $Points
        if (-not [string]::IsNullOrWhiteSpace($Text)) {
            $Bucket.Evidence.Add($Text)
        }
    }

    function Add-HypothesisCounterEvidence {
        param(
            $Bucket,
            [string]$Text
        )

        if (-not [string]::IsNullOrWhiteSpace($Text)) {
            $Bucket.CounterEvidence.Add($Text)
        }
    }

    $findings = New-Object System.Collections.Generic.List[string]
    $discarded = New-Object System.Collections.Generic.List[string]
    $limitations = New-Object System.Collections.Generic.List[string]
    $signals = New-Object System.Collections.Generic.List[string]
    $inconclusive = New-Object System.Collections.Generic.List[string]

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

    $requestedSampleTarget = $sampleCount
    $timingIntegrity = 'OK'
    $timingNote = 'Janela real respeitou o orçamento temporal configurado.'
    $targetWallClockSeconds = $null
    $actualWallClockSeconds = $null
    $actualWallClockMinutes = $null
    $timingDriftSeconds = $null
    $stoppedByWallClock = $false

    if ($Timeline.PSObject.Properties.Name -contains 'RequestedSampleTarget') {
        try { $requestedSampleTarget = [int]$Timeline.RequestedSampleTarget } catch {}
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TimingIntegrity') {
        $timingIntegrity = [string]$Timeline.TimingIntegrity
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TimingNote') {
        $timingNote = [string]$Timeline.TimingNote
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TargetWallClockSeconds') {
        $targetWallClockSeconds = $Timeline.TargetWallClockSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'ActualWallClockSeconds') {
        $actualWallClockSeconds = $Timeline.ActualWallClockSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'ActualWallClockMinutes') {
        $actualWallClockMinutes = $Timeline.ActualWallClockMinutes
    }
    if ($Timeline.PSObject.Properties.Name -contains 'TimingDriftSeconds') {
        $timingDriftSeconds = $Timeline.TimingDriftSeconds
    }
    if ($Timeline.PSObject.Properties.Name -contains 'StoppedByWallClock') {
        $stoppedByWallClock = [bool]$Timeline.StoppedByWallClock
    }

    $dbUnavailableRatio = [math]::Round(($dbUnavailableSamples / [double]$sampleCount), 2)
    $netUnavailableRatio = [math]::Round(($netUnavailableSamples / [double]$sampleCount), 2)

    $hypotheses = [ordered]@{
        SHARE = (New-HypothesisBucket -Key 'SHARE' -Display 'Compartilhamento/acesso' -ImpactScope 'Sistema compartilhado' -RecommendedAction 'Validar disponibilidade do caminho do banco/NetDir, latência do compartilhamento e permissões antes de atuar no software.')
        LOCAL = (New-HypothesisBucket -Key 'LOCAL' -Display 'Configuração local' -ImpactScope 'Somente este computador' -RecommendedAction 'Conferir configuração local do ECG/BDE, mapeamentos e aderência ao contrato da unidade.')
        LOCK = (New-HypothesisBucket -Key 'LOCK' -Display 'Contenção/lock' -ImpactScope 'Sistema compartilhado' -RecommendedAction 'Repetir a rodada durante o sintoma e revisar concorrência de acesso, locks e fluxo de gravação no NetDir.')
        SOFTWARE = (New-HypothesisBucket -Key 'SOFTWARE' -Display 'Software/arquivo' -ImpactScope 'Somente este computador' -RecommendedAction 'Revisar saúde do software do ECG e comportamento local da estação, priorizando logs e estado do aplicativo.')
    }

    if ($Paths.PSObject.Properties.Name -contains 'BdeNetDirStatus') {
        switch ([string]$Paths.BdeNetDirStatus) {
            'AUSENTE' {
                $findings.Add('NETDIR do BDE está ausente no registro da estação.')
                $signals.Add('BDE_NETDIR_MISSING')
                Add-HypothesisEvidence -Bucket $hypotheses.LOCAL -Points 5 -Text 'NETDIR do BDE ausente no registro local.'
            }
            'DIVERGENTE' {
                $findings.Add('NETDIR do BDE diverge do caminho esperado para a unidade.')
                $signals.Add('BDE_NETDIR_DIVERGENT')
                Add-HypothesisEvidence -Bucket $hypotheses.LOCAL -Points 5 -Text 'NETDIR do BDE divergente do contrato esperado.'
            }
            'OK' {
                $discarded.Add('NETDIR do BDE está aderente ao contrato esperado para a unidade.')
                Add-HypothesisCounterEvidence -Bucket $hypotheses.LOCAL -Text 'NETDIR do BDE aderente ao contrato esperado.'
            }
        }
    }

    if ($Paths.PSObject.Properties.Name -contains 'LockControlFilePresent') {
        if (($Paths.NetDirAccessible -eq $true) -and ($Paths.LockControlFilePresent -eq $false)) {
            $findings.Add('NetDir acessível, porém o arquivo PDOXUSRS.NET não foi encontrado.')
            $signals.Add('BDE_PDOXUSRS_NET_MISSING')
            Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points 3 -Text 'Arquivo PDOXUSRS.NET ausente no NetDir acessível.'
        }
        elseif ($Paths.LockControlFilePresent -eq $true) {
            $discarded.Add('Arquivo PDOXUSRS.NET presente em um dos caminhos válidos de NetDir.')
            Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'Arquivo PDOXUSRS.NET presente em caminho válido de NetDir.'
        }
    }

    if (-not $Paths.ExeAccessible) {
        $findings.Add('Executável oficial do ECG não foi localizado no caminho padrão da ferramenta.')
        $signals.Add('EXE_OFFICIAL_NOT_FOUND')
        Add-HypothesisEvidence -Bucket $hypotheses.LOCAL -Points 4 -Text 'Executável oficial inacessível na estação durante a rodada.'
    }
    else {
        $discarded.Add('Executável oficial do ECG localizado no caminho padrão.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.LOCAL -Text 'Executável oficial acessível no caminho padrão.'
    }

    if (-not $Paths.DatabaseAccessible) {
        $findings.Add('Banco do ECG inacessível no caminho efetivo da rodada.')
        $signals.Add('DATABASE_PATH_UNAVAILABLE')
        Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points 5 -Text 'Banco inacessível no caminho efetivo durante a rodada.'
    }
    else {
        $discarded.Add('Banco do ECG acessível no caminho efetivo da rodada.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'Banco acessível no caminho efetivo nesta rodada.'
    }

    if (-not $Paths.NetDirAccessible) {
        $findings.Add('NetDir inacessível no caminho efetivo da rodada.')
        $signals.Add('NETDIR_PATH_UNAVAILABLE')
        Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points 5 -Text 'NetDir inacessível no caminho efetivo durante a rodada.'
    }
    else {
        $discarded.Add('NetDir acessível no caminho efetivo da rodada.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'NetDir acessível no caminho efetivo nesta rodada.'
    }

    if ([string]$Paths.DatabasePathSource -match 'não acessível|Fallback') {
        $findings.Add('O caminho do banco foi resolvido apenas por fallback lógico, sem acessibilidade plena nesta rodada.')
        $signals.Add('DATABASE_SOURCE_FALLBACK')
        Add-HypothesisEvidence -Bucket $hypotheses.LOCAL -Points 3 -Text 'Caminho do banco caiu em fallback lógico ou inacessível nesta rodada.'
    }
    else {
        $discarded.Add('Banco resolvido de forma consistente para o perfil e/ou configuração atual da unidade.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.LOCAL -Text 'Banco resolvido de forma consistente para o perfil e/ou configuração atual da unidade.'
    }

    if ($dbUnavailableSamples -gt 0) {
        $findings.Add('Houve indisponibilidade do banco em ' + $dbUnavailableSamples + ' amostra(s) da rodada.')
        $signals.Add('DATABASE_ACCESS_INSTABILITY')
        Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points ([Math]::Min(4, [Math]::Max(1, $dbUnavailableSamples))) -Text ('Banco instável em ' + $dbUnavailableSamples + ' amostra(s) da rodada.')
    }
    else {
        $discarded.Add('Sem indisponibilidade observada do banco durante a janela coletada.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'Sem indisponibilidade de banco na janela observada.'
    }

    if ($netUnavailableSamples -gt 0) {
        $findings.Add('Houve indisponibilidade do NetDir em ' + $netUnavailableSamples + ' amostra(s) da rodada.')
        $signals.Add('NETDIR_ACCESS_INSTABILITY')
        Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points ([Math]::Min(4, [Math]::Max(1, $netUnavailableSamples))) -Text ('NetDir instável em ' + $netUnavailableSamples + ' amostra(s) da rodada.')
    }
    else {
        $discarded.Add('Sem indisponibilidade observada do NetDir durante a janela coletada.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'Sem indisponibilidade de NetDir na janela observada.'
    }

    if ($null -ne $peakLocks -and $peakLocks -ge 1) {
        $findings.Add('Arquivos de lock/controle foram observados no NetDir durante a janela (pico: ' + $peakLocks + ').')
        $signals.Add('LOCK_ACTIVITY_OBSERVED')
        Add-HypothesisEvidence -Bucket $hypotheses.LOCK -Points ([Math]::Min(5, [Math]::Max(2, $peakLocks + 1))) -Text ('Atividade de lock/controle observada no NetDir (pico: ' + $peakLocks + ').')
    }
    else {
        $discarded.Add('Sem lock/controle relevante observado no NetDir durante a janela.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.LOCK -Text 'Sem lock/controle relevante no NetDir durante a janela observada.'
    }

    if ($null -ne $avgCpu) {
        if ($avgCpu -ge 70) {
            $findings.Add('CPU média da rodada em faixa elevada (' + $avgCpu + '%).')
            $signals.Add('CPU_AVERAGE_HIGH')
            Add-HypothesisEvidence -Bucket $hypotheses.SOFTWARE -Points 3 -Text ('CPU média elevada na rodada (' + $avgCpu + '%).')
        }
        elseif ($avgCpu -lt 55) {
            $discarded.Add('CPU média da rodada sem pressão sustentada relevante (' + $avgCpu + '%).')
            Add-HypothesisCounterEvidence -Bucket $hypotheses.SOFTWARE -Text ('CPU média sem pressão sustentada relevante (' + $avgCpu + '%).')
        }
        else {
            $inconclusive.Add('CPU média ficou em faixa intermediária; o sinal isolado não fecha hipótese por si só.')
        }
    }
    else {
        $inconclusive.Add('CPU média indisponível nesta rodada.')
    }

    if ($null -ne $peakCpu) {
        if ($peakCpu -ge 90) {
            $findings.Add('Pico de CPU relevante durante a rodada (' + $peakCpu + '%).')
            $signals.Add('CPU_PEAK_HIGH')
            Add-HypothesisEvidence -Bucket $hypotheses.SOFTWARE -Points 2 -Text ('Pico de CPU relevante observado (' + $peakCpu + '%).')
        }
        elseif ($peakCpu -lt 80) {
            $discarded.Add('Sem pico extremo de CPU durante a rodada (' + $peakCpu + '%).')
            Add-HypothesisCounterEvidence -Bucket $hypotheses.SOFTWARE -Text ('Sem pico extremo de CPU durante a rodada (' + $peakCpu + '%).')
        }
        else {
            $inconclusive.Add('Houve pico de CPU moderado, mas sem pressão extrema sustentada.')
        }
    }
    else {
        $inconclusive.Add('Pico de CPU indisponível nesta rodada.')
    }

    if ($smbTimeoutSamples -gt 0) {
        $findings.Add('Parte das consultas SMB excedeu timeout; leitura de compartilhamento deve ser interpretada com cautela.')
        $signals.Add('SMB_TIMEOUTS')
        Add-HypothesisEvidence -Bucket $hypotheses.SHARE -Points ([Math]::Min(3, [Math]::Max(1, $smbTimeoutSamples))) -Text ('Consultas SMB excederam timeout em ' + $smbTimeoutSamples + ' amostra(s).')
        $inconclusive.Add('Consultas SMB com timeout reduzem a força de alguns descartes de compartilhamento.')
    }
    else {
        $discarded.Add('Sem timeout observado nas consultas SMB executadas pela ferramenta.')
        Add-HypothesisCounterEvidence -Bucket $hypotheses.SHARE -Text 'Sem timeout observado nas consultas SMB executadas pela ferramenta.'
    }

    if (($hypotheses.SHARE.Score -eq 0) -and ($hypotheses.LOCAL.Score -eq 0) -and ($hypotheses.LOCK.Score -eq 0) -and ($hypotheses.SOFTWARE.Score -eq 0)) {
        $inconclusive.Add('A rodada não produziu sinal dominante suficiente para priorizar uma hipótese acima das demais.')
    }

    if ($timingIntegrity -ne 'OK') {
        $actualText = if ($null -ne $actualWallClockSeconds) { ([string]$actualWallClockSeconds) + ' s' } else { 'N/D' }
        $targetText = if ($null -ne $targetWallClockSeconds) { ([string]$targetWallClockSeconds) + ' s' } else { 'N/D' }
        $findings.Add('Janela temporal da coleta saiu do orçamento previsto (' + $actualText + ' reais vs ' + $targetText + ' previstos).')
        $signals.Add('TIMING_WINDOW_DRIFT')
        $limitations.Add('Integridade temporal da coleta: ' + $timingIntegrity + '. ' + $timingNote)
        $inconclusive.Add('Deriva temporal reduz comparabilidade entre rodadas e sugere custo variável por amostra.')
        if ($stoppedByWallClock -eq $true) {
            $inconclusive.Add('Watchdog temporal encerrou a coleta antes de atingir a meta ideal de amostras para preservar a janela real configurada.')
        }
    }
    elseif ($requestedSampleTarget -ne $sampleCount) {
        $limitations.Add('Meta ideal de amostras não foi atingida nesta rodada, mesmo com janela temporal íntegra.')
    }

    $limitations.Add('Esta versão não executa benchmark assistido dentro do ECG; usa índice observacional da rodada.')
    $limitations.Add('Esta versão não executa comparação com referência.')
    $limitations.Add('Esta versão não executa avaliação Defender/minifilter.')

    $rankedHypotheses = @(
        $hypotheses.SHARE,
        $hypotheses.LOCAL,
        $hypotheses.LOCK,
        $hypotheses.SOFTWARE
    ) | Sort-Object -Property @{ Expression = { [int]$_.Score }; Descending = $true }, @{ Expression = { [string]$_.Display }; Descending = $false }

    $primaryHypothesisBucket = $rankedHypotheses[0]
    $secondaryHypothesisBuckets = @()
    if ($rankedHypotheses.Count -gt 1) {
        $secondaryHypothesisBuckets = @($rankedHypotheses | Select-Object -Skip 1)
    }

    $primaryHypothesis = 'Ainda sem hipótese principal definida'
    $recommendedAction = 'Executar nova rodada durante o sintoma e cruzar com observação do operador para consolidar a hipótese.'
    $impactScope = 'Ainda não foi possível definir'

    if ($primaryHypothesisBucket.Score -gt 0) {
        $primaryHypothesis = $primaryHypothesisBucket.Display
        $recommendedAction = $primaryHypothesisBucket.RecommendedAction
        $impactScope = $primaryHypothesisBucket.ImpactScope
    }
    elseif ($MachineInfo.MachineType -eq 'Servidor de arquivos') {
        $impactScope = 'Sistema compartilhado'
    }

    $status = 'INCONCLUSIVO'
    if ((-not $Paths.ExeAccessible) -or ($dbUnavailableRatio -ge 0.30) -or ($netUnavailableRatio -ge 0.30)) {
        $status = 'CRÍTICO'
    }
    elseif (($severityScore -ge 30) -or (($null -ne $peakLocks) -and ($peakLocks -ge 1)) -or (($null -ne $avgCpu) -and ($avgCpu -ge 65)) -or ($smbTimeoutSamples -gt 0)) {
        $status = 'LENTO'
    }
    elseif ($severityScore -eq 0 -and $Paths.ExeAccessible -and $Paths.DatabaseAccessible -and $Paths.NetDirAccessible) {
        $status = 'NORMAL'
    }

    $confidence = 'Baixa'
    $topScore = [int]$primaryHypothesisBucket.Score
    $topEvidenceCount = @($primaryHypothesisBucket.Evidence).Count
    $secondScore = 0
    if ($rankedHypotheses.Count -gt 1) {
        $secondScore = [int]$rankedHypotheses[1].Score
    }
    $scoreGap = $topScore - $secondScore

    if ($status -eq 'NORMAL' -and $topScore -eq 0) {
        $confidence = 'Média'
    }
    elseif ($topScore -ge 8 -and $topEvidenceCount -ge 2 -and $scoreGap -ge 3) {
        $confidence = 'Alta'
    }
    elseif ($topScore -ge 4 -and $topEvidenceCount -ge 1 -and $scoreGap -ge 1) {
        $confidence = 'Média'
    }

    if ($timingIntegrity -eq 'Crítica') {
        $confidence = 'Baixa'
    }
    elseif (($timingIntegrity -eq 'Atenção') -and ($confidence -eq 'Alta')) {
        $confidence = 'Média'
    }

    $probablePerception = 'Usuário percebe lentidão ou intermitência operacional no ECG, especialmente na etapa priorizada.'
    if ($status -eq 'NORMAL') {
        $probablePerception = 'Usuário pode ter percebido oscilação anterior, mas a rodada atual não reuniu evidência forte de degradação ativa.'
    }
    elseif ($status -eq 'CRÍTICO') {
        $probablePerception = 'Usuário tende a perceber falha evidente, demora anormal ou impossibilidade prática de continuar o fluxo do ECG.'
    }

    $summaryPhrase = 'A rodada não reuniu evidência suficiente para fechar hipótese principal com segurança.'
    if ($status -eq 'NORMAL') {
        $summaryPhrase = 'A rodada foi concluída sem evidência forte de degradação ativa nas camadas observadas pela ferramenta.'
    }
    elseif ($primaryHypothesisBucket.Score -gt 0) {
        if ($status -eq 'CRÍTICO') {
            $summaryPhrase = 'A rodada indica condição crítica com hipótese principal em ' + $primaryHypothesis.ToLowerInvariant() + '.'
        }
        else {
            $summaryPhrase = 'A rodada indica degradação operacional com hipótese principal em ' + $primaryHypothesis.ToLowerInvariant() + '.'
        }
    }

    if ($findings.Count -eq 0) {
        $findings.Add('A rodada não encontrou evidência forte suficiente para afirmar uma hipótese principal acima das demais.')
    }
    if ($discarded.Count -eq 0) {
        $discarded.Add('A rodada ainda não produziu descartes técnicos fortes.')
    }
    if ($inconclusive.Count -eq 0) {
        $inconclusive.Add('Não houve lacuna adicional relevante além das limitações já declaradas pela ferramenta.')
    }

    $secondaryHypothesisObjects = @()
    foreach ($bucket in $secondaryHypothesisBuckets) {
        $secondaryHypothesisObjects += [PSCustomObject]@{
            Name = $bucket.Display
            Score = [int]$bucket.Score
            MainReasons = @($bucket.Evidence | Select-Object -First 2)
        }
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
        RequestedSampleTarget = $requestedSampleTarget
        TimingIntegrity = $timingIntegrity
        TimingNote = $timingNote
        TargetWallClockSeconds = $targetWallClockSeconds
        ActualWallClockSeconds = $actualWallClockSeconds
        ActualWallClockMinutes = $actualWallClockMinutes
        TimingDriftSeconds = $timingDriftSeconds
        StoppedByWallClock = $stoppedByWallClock
        Status = $statusDisplay
        StatusCode = $status
        Confidence = $confidenceDisplay
        PrimaryHypothesis = $primaryHypothesis
        ProbableCause = $primaryHypothesis
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
        InconclusivePoints = $inconclusive.ToArray()
        HypothesisSupport = @($primaryHypothesisBucket.Evidence)
        HypothesisCounterpoints = @($primaryHypothesisBucket.CounterEvidence)
        SecondaryHypotheses = @($secondaryHypothesisObjects)
        PassiveBenchmarkSummary = $PassiveBenchmark.SummaryPhrase
        Metrics = [PSCustomObject]@{
            SampleCount = $Timeline.SampleCount
            RequestedSampleTarget = $requestedSampleTarget
            TargetWallClockSeconds = $targetWallClockSeconds
            ActualWallClockSeconds = $actualWallClockSeconds
            ActualWallClockMinutes = $actualWallClockMinutes
            TimingDriftSeconds = $timingDriftSeconds
            TimingIntegrity = $timingIntegrity
            StoppedByWallClock = $stoppedByWallClock
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
            TopHypothesisScore = $topScore
            SecondHypothesisScore = $secondScore
            HypothesisScoreGap = $scoreGap
        }
    }
}



function Build-SummaryText {
    param($Analysis)

    $supportLines = @()
    foreach ($line in @($Analysis.HypothesisSupport)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $supportLines += ('- ' + [string]$line)
        }
    }
    if ($supportLines.Count -eq 0) {
        $supportLines = @('- Ainda sem evidência dominante suficiente para sustentar uma hipótese principal forte.')
    }

    $counterLines = @()
    foreach ($line in @($Analysis.HypothesisCounterpoints)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $counterLines += ('- ' + [string]$line)
        }
    }
    if ($counterLines.Count -eq 0) {
        $counterLines = @('- Ainda sem contrapontos fortes contra hipóteses rivais nesta rodada.')
    }

    $inconclusiveLines = @()
    foreach ($line in @($Analysis.InconclusivePoints)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $inconclusiveLines += ('- ' + [string]$line)
        }
    }
    if ($inconclusiveLines.Count -eq 0) {
        $inconclusiveLines = @('- Sem ponto inconclusivo adicional relevante nesta rodada.')
    }

    $discardLines = @()
    foreach ($line in @($Analysis.WhatDidNotIndicateFailure)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $discardLines += ('- ' + [string]$line)
        }
    }
    if ($discardLines.Count -eq 0) {
        $discardLines = @('- Ainda sem descarte confiável nesta rodada.')
    }

    $timingLines = @()
    if ($Analysis.PSObject.Properties.Name -contains 'ActualWallClockMinutes' -and $null -ne $Analysis.ActualWallClockMinutes) {
        $timingLines += ('Janela real (relógio): ' + [string]$Analysis.ActualWallClockMinutes + ' minuto(s)')
    }
    if ($Analysis.PSObject.Properties.Name -contains 'TimingIntegrity' -and -not [string]::IsNullOrWhiteSpace([string]$Analysis.TimingIntegrity)) {
        $timingLines += ('Integridade temporal: ' + [string]$Analysis.TimingIntegrity)
    }
    $timingText = ''
    if ($timingLines.Count -gt 0) {
        $timingText = [Environment]::NewLine + ($timingLines -join [Environment]::NewLine)
    }

@"
Data/hora da coleta: $($Analysis.CollectedAt)
Máquina analisada: $($Analysis.ComputerName) — $($Analysis.MachineType)
Etapa priorizada: $($Analysis.StageLabel)
Sintoma informado: $($Analysis.SymptomLabel)
Janela de observação: $($Analysis.ObservationMinutes) minuto(s)$timingText

Status: $($Analysis.Status)
Hipótese principal desta rodada: $($Analysis.PrimaryHypothesis)
Confiança: $($Analysis.Confidence)
Alcance provável: $($Analysis.ImpactScope)

Resumo executivo:
$($Analysis.SummaryPhrase)

Evidências que sustentam a hipótese principal:
$($supportLines -join [Environment]::NewLine)

O que enfraqueceu hipóteses rivais:
$($counterLines -join [Environment]::NewLine)

Pontos ainda inconclusivos:
$($inconclusiveLines -join [Environment]::NewLine)

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

    $normalizedSeries = @()
    foreach ($item in @($series)) {
        $normalizedSeries += [PSCustomObject]@{
            name = $item.Name
            values = @($item.Values)
            max = $item.Max
            color = $item.Color
        }
    }

    return [PSCustomObject]@{
        labels = @($labels)
        series = @($normalizedSeries)
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

    $supportItems = ''
    foreach ($line in @($Analysis.HypothesisSupport)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $supportItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
        }
    }
    if ([string]::IsNullOrWhiteSpace($supportItems)) {
        $supportItems = '<li>Ainda sem evidência dominante suficiente para sustentar uma hipótese principal forte.</li>'
    }

    $counterItems = ''
    foreach ($line in @($Analysis.HypothesisCounterpoints)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$line)) {
            $counterItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
        }
    }
    if ([string]::IsNullOrWhiteSpace($counterItems)) {
        $counterItems = '<li>Ainda sem contrapontos fortes contra hipóteses rivais nesta rodada.</li>'
    }

    $findingItems = ''
    foreach ($line in @($Analysis.Findings)) {
        $findingItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($findingItems)) {
        $findingItems = '<li>Sem achado principal adicional nesta rodada.</li>'
    }

    $discardItems = ''
    foreach ($line in @($Analysis.WhatDidNotIndicateFailure)) {
        $discardItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($discardItems)) {
        $discardItems = '<li>Ainda sem descarte confiável nesta rodada.</li>'
    }

    $inconclusiveItems = ''
    foreach ($line in @($Analysis.InconclusivePoints)) {
        $inconclusiveItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($inconclusiveItems)) {
        $inconclusiveItems = '<li>Sem ponto inconclusivo adicional relevante nesta rodada.</li>'
    }

    $limitationItems = ''
    foreach ($line in @($Analysis.Limitations)) {
        $limitationItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($limitationItems)) {
        $limitationItems = '<li>Sem limitação declarada adicional.</li>'
    }

    $pathNoteItems = ''
    foreach ($line in @($Paths.Notes)) {
        $pathNoteItems += '<li>' + (HtmlEncode ([string]$line)) + '</li>'
    }
    if ([string]::IsNullOrWhiteSpace($pathNoteItems)) {
        $pathNoteItems = '<li>Sem observação adicional de resolução de paths nesta rodada.</li>'
    }

    $secondaryHypothesisRows = ''
    foreach ($item in @($Analysis.SecondaryHypotheses)) {
        $reasons = ''
        foreach ($reason in @($item.MainReasons)) {
            if (-not [string]::IsNullOrWhiteSpace([string]$reason)) {
                if (-not [string]::IsNullOrWhiteSpace($reasons)) {
                    $reasons += '; '
                }
                $reasons += [string]$reason
            }
        }
        if ([string]::IsNullOrWhiteSpace($reasons)) {
            $reasons = 'Sem evidência relevante nesta rodada.'
        }

        $secondaryHypothesisRows += '<tr>' +
            '<td>' + (HtmlEncode ([string]$item.Name)) + '</td>' +
            '<td>' + (HtmlEncode ([string]$item.Score)) + '</td>' +
            '<td>' + (HtmlEncode $reasons) + '</td>' +
            '</tr>'
    }
    if ([string]::IsNullOrWhiteSpace($secondaryHypothesisRows)) {
        $secondaryHypothesisRows = '<tr><td colspan="3">Sem hipótese secundária relevante nesta rodada.</td></tr>'
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
        $dbOkCell = if ($sample.DatabaseAccessible -eq $true) { 'Sim' } else { 'Não' }
        $netOkCell = if ($sample.NetDirAccessible -eq $true) { 'Sim' } else { 'Não' }

        $timelineRows += '<tr>' +
            '<td>' + (HtmlEncode ([string]$sample.Timestamp)) + '</td>' +
            '<td>' + (HtmlEncode $cpuCell) + '</td>' +
            '<td>' + (HtmlEncode $lockCell) + '</td>' +
            '<td>' + (HtmlEncode $connCell) + '</td>' +
            '<td>' + (HtmlEncode $sessCell) + '</td>' +
            '<td>' + (HtmlEncode $openCell) + '</td>' +
            '<td>' + (HtmlEncode $openNetCell) + '</td>' +
            '<td>' + (HtmlEncode $dbOkCell) + '</td>' +
            '<td>' + (HtmlEncode $netOkCell) + '</td>' +
            '<td>' + (HtmlEncode $timeoutCell) + '</td>' +
            '</tr>'
    }
    if ([string]::IsNullOrWhiteSpace($timelineRows)) {
        $timelineRows = '<tr><td colspan="10">Sem amostra disponível na timeline desta rodada.</td></tr>'
    }

    $chartDefinition = Get-ChartDefinition -Timeline $Timeline
    $chartJson = $chartDefinition | ConvertTo-Json -Depth 10 -Compress

    $summaryJsonPreview = [ordered]@{
        Status = $Analysis.Status
        Confidence = $Analysis.Confidence
        PrimaryHypothesis = $Analysis.PrimaryHypothesis
        ImpactScope = $Analysis.ImpactScope
        SummaryPhrase = $Analysis.SummaryPhrase
        RecommendedAction = $Analysis.RecommendedAction
        HypothesisSupport = @($Analysis.HypothesisSupport)
        InconclusivePoints = @($Analysis.InconclusivePoints)
    }
    $summaryJsonText = $summaryJsonPreview | ConvertTo-Json -Depth 5

    $timingActualMinutesText = 'N/D'
    if ($PassiveBenchmark.PSObject.Properties.Name -contains 'ActualWallClockMinutes' -and $null -ne $PassiveBenchmark.ActualWallClockMinutes) {
        $timingActualMinutesText = [string]$PassiveBenchmark.ActualWallClockMinutes
    }

    $timingIntegrityText = 'N/D'
    if ($PassiveBenchmark.PSObject.Properties.Name -contains 'TimingIntegrity' -and -not [string]::IsNullOrWhiteSpace([string]$PassiveBenchmark.TimingIntegrity)) {
        $timingIntegrityText = [string]$PassiveBenchmark.TimingIntegrity
    }

    $timingDriftText = 'N/D'
    if ($PassiveBenchmark.PSObject.Properties.Name -contains 'TimingDriftSeconds' -and $null -ne $PassiveBenchmark.TimingDriftSeconds) {
        $timingDriftText = [string]$PassiveBenchmark.TimingDriftSeconds + ' s'
    }

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
.grid-3 { display: grid; gap: 12px; grid-template-columns: repeat(3, minmax(0, 1fr)); }
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
.pre-summary { white-space: pre-wrap; word-break: break-word; background: #0f172a; color: #e2e8f0; padding: 14px; border-radius: 10px; overflow-x: auto; }
.highlight { border: 1px solid #dbeafe; background: #f8fbff; }
@media print {
  body { background: #ffffff; }
  .wrapper { max-width: none; padding: 0; }
  .card, details { box-shadow: none; border: 1px solid #d1d5db; }
}
@media (max-width: 960px) {
  .grid, .grid-3 { grid-template-columns: 1fr; }
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
            <div class="kv"><strong>Hipótese principal desta rodada</strong>$([string](HtmlEncode $Analysis.PrimaryHypothesis))</div>
            <div class="kv"><strong>Confiança</strong>$([string](HtmlEncode $Analysis.Confidence))</div>
            <div class="kv"><strong>Tempo de observação</strong>$([string](HtmlEncode ([string]$Analysis.ObservationMinutes))) minuto(s)</div>
            <div class="kv"><strong>Alcance provável</strong>$([string](HtmlEncode $Analysis.ImpactScope))</div>
            <div class="kv"><strong>RunId</strong>$([string](HtmlEncode $Analysis.RunId))</div>
            <div class="kv"><strong>Executado por</strong>$([string](HtmlEncode $Analysis.ExecutedBy))</div>
        </div>
    </div>

    <div class="card highlight">
        <h2>Diagnóstico executivo</h2>
        <p><strong>Leitura operacional:</strong> $([string](HtmlEncode $Analysis.ProbablePerception))</p>
        <p><strong>Próxima ação recomendada:</strong> $([string](HtmlEncode $Analysis.RecommendedAction))</p>
        <div class="note">
            <strong>Uso correto deste laudo:</strong> a conclusão prioriza a melhor hipótese desta rodada com base nos sinais observados, sem vender certeza acima da evidência disponível.
        </div>
    </div>

    <div class="grid">
        <div class="card">
            <h2>Evidências que sustentam a hipótese principal</h2>
            <ul>
                $supportItems
            </ul>
        </div>
        <div class="card">
            <h2>O que enfraqueceu hipóteses rivais</h2>
            <ul>
                $counterItems
            </ul>
        </div>
    </div>

    <div class="grid">
        <div class="card">
            <h2>O que não indicou falha relevante</h2>
            <ul>
                $discardItems
            </ul>
        </div>
        <div class="card">
            <h2>Pontos ainda inconclusivos</h2>
            <ul>
                $inconclusiveItems
            </ul>
        </div>
    </div>

    <div class="card">
        <h2>Hipóteses secundárias da rodada</h2>
        <table>
            <tr>
                <th>Hipótese</th>
                <th>Score</th>
                <th>Principais razões</th>
            </tr>
            $secondaryHypothesisRows
        </table>
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
                    <div class="kv"><strong>Meta ideal de amostras</strong>$([string](HtmlEncode ([string]$PassiveBenchmark.RequestedSampleTarget)))</div>
                    <div class="kv"><strong>Janela real (relógio)</strong>$([string](HtmlEncode $timingActualMinutesText)) min</div>
                    <div class="kv"><strong>Integridade temporal</strong>$([string](HtmlEncode $timingIntegrityText))</div>
                    <div class="kv"><strong>Drift temporal</strong>$([string](HtmlEncode $timingDriftText))</div>
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
                    <tr><td>Unidade detectada</td><td>$([string](HtmlEncode $MachineInfo.UnitCode))</td></tr>
                    <tr><td>Perfil operacional</td><td>$([string](HtmlEncode $MachineInfo.ProfileName))</td></tr>
                    <tr><td>Topologia</td><td>$([string](HtmlEncode $MachineInfo.TopologyType))</td></tr>
                    <tr><td>Tipo da máquina</td><td>$([string](HtmlEncode $MachineInfo.MachineType))</td></tr>
                    <tr><td>Usuário esperado</td><td>$([string](HtmlEncode $MachineInfo.ExpectedUser))</td></tr>
                    <tr><td>Usuário esperado confere</td><td>$([string](HtmlEncode ([string]$MachineInfo.ExpectedUserMatch)))</td></tr>
                    <tr><td>Executável oficial</td><td>$([string](HtmlEncode $Paths.ExePath))</td></tr>
                    <tr><td>Executável acessível</td><td>$([string](HtmlEncode ([string]$Paths.ExeAccessible)))</td></tr>
                    <tr><td>Banco esperado</td><td>$([string](HtmlEncode $Paths.DatabasePathExpected))</td></tr>
                    <tr><td>Banco efetivo</td><td>$([string](HtmlEncode $Paths.DatabasePath))</td></tr>
                    <tr><td>Origem do banco</td><td>$([string](HtmlEncode $Paths.DatabasePathSource))</td></tr>
                    <tr><td>Banco acessível</td><td>$([string](HtmlEncode ([string]$Paths.DatabaseAccessible)))</td></tr>
                    <tr><td>NETDIR esperado</td><td>$([string](HtmlEncode $Paths.NetDirPathExpected))</td></tr>
                    <tr><td>NETDIR atual no BDE</td><td>$([string](HtmlEncode $Paths.CurrentBdeNetDir))</td></tr>
                    <tr><td>Status do NETDIR</td><td>$([string](HtmlEncode $Paths.BdeNetDirStatus))</td></tr>
                    <tr><td>NetDir efetivo</td><td>$([string](HtmlEncode $Paths.NetDirPath))</td></tr>
                    <tr><td>Origem do NetDir</td><td>$([string](HtmlEncode $Paths.NetDirSource))</td></tr>
                    <tr><td>NetDir acessível</td><td>$([string](HtmlEncode ([string]$Paths.NetDirAccessible)))</td></tr>
                    <tr><td>PDOXUSRS.NET presente</td><td>$([string](HtmlEncode ([string]$Paths.LockControlFilePresent)))</td></tr>
                </table>
                <h4>Observações de resolução de paths</h4>
                <ul>
                    $pathNoteItems
                </ul>
            </div>

            <div class="card" style="box-shadow:none; border:1px solid #e5e7eb;">
                <h3>Limitações desta rodada</h3>
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
                <div class="pre-summary">$([string](HtmlEncode $summaryJsonText))</div>
            </div>
        </div>
    </details>
</div>
<script>
(function () {
    var chartData = $chartJson || {};
    chartData.labels = chartData.labels || chartData.Labels || [];
    chartData.series = chartData.series || chartData.Series || [];
    chartData.series = chartData.series.map(function (series) {
        return {
            name: series.name || series.Name || '',
            values: series.values || series.Values || [],
            max: (series.max !== undefined && series.max !== null) ? series.max : series.Max,
            color: series.color || series.Color || '#2563eb'
        };
    });
    if (!chartData.series || chartData.series.length === 0) {
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
    $paths = Resolve-EcgPaths -MachineInfo $machineInfo

    $context = [PSCustomObject]@{
        ToolName = $ToolName
        ToolVersion = $ToolVersion
        RunId = $script:RunId
        CollectedAt = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
        Machine = $machineInfo
        UnitProfileFile = $UnitProfilesFile
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
