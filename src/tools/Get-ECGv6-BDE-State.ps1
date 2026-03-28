<#
.SYNOPSIS
    Coleta informações do ambiente ECG/BDE para diagnóstico.
.DESCRIPTION
    Gera relatórios JSON e TXT com dados de sistema, registros, caminhos, permissões, etc.
    Se a política de execução estiver restrita, o script se reexecuta automaticamente com Bypass.
.PARAMETER OutDir
    Diretório onde os arquivos de saída serão salvos (padrão: C:\ECG\Diag).
.PARAMETER IncludeAcl
    Inclui ACLs na inspeção de caminhos.
.PARAMETER WriteProbe
    Testa permissão de escrita em cada caminho inspecionado.
.PARAMETER CreateMissingDirs
    Cria os diretórios C:\HW\NetDir e C:\HW\Private se não existirem.
.EXAMPLE
    .\Get-ECGv6-BDE-State.ps1 -IncludeAcl -WriteProbe -CreateMissingDirs
#>

[CmdletBinding()]
param(
    [string]$OutDir = 'C:\ECG\Diag',
    [switch]$IncludeAcl,
    [switch]$WriteProbe,
    [switch]$CreateMissingDirs
)

# --- Auto-elevação e bypass de política ---
# Se a política de execução for Restricted, reexecuta com Bypass
if ($PSVersionTable.PSVersion.Major -ge 3) {
    $execPolicy = Get-ExecutionPolicy -Scope CurrentUser
    if ($execPolicy -eq 'Restricted') {
        Write-Host "Política de execução é 'Restricted'. Reiniciando com ExecutionPolicy Bypass..." -ForegroundColor Yellow
        $scriptPath = $MyInvocation.MyCommand.Path
        $arguments = @("-ExecutionPolicy", "Bypass", "-File", "`"$scriptPath`"")
        if ($IncludeAcl) { $arguments += "-IncludeAcl" }
        if ($WriteProbe) { $arguments += "-WriteProbe" }
        if ($CreateMissingDirs) { $arguments += "-CreateMissingDirs" }
        if ($OutDir) { $arguments += "-OutDir", "`"$OutDir`"" }
        Start-Process powershell -ArgumentList $arguments -Wait -NoNewWindow
        exit
    }
}

$ErrorActionPreference = 'Stop'

# --- Criação opcional de diretórios ausentes ---
if ($CreateMissingDirs) {
    $dirsToCreate = @('C:\HW\NetDir', 'C:\HW\Private')
    foreach ($dir in $dirsToCreate) {
        if (-not (Test-Path $dir)) {
            try {
                New-Item -ItemType Directory -Path $dir -Force | Out-Null
                Write-Host "Diretório criado: $dir" -ForegroundColor Green
            }
            catch {
                Write-Host "Falha ao criar $dir : $($_.Exception.Message)" -ForegroundColor Red
            }
        }
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

function Normalize-PathString {
    param([string]$Path)

    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }

    $p = ($Path -replace '/', '\').Trim()

    while ($p.Length -gt 3 -and $p.EndsWith('\')) {
        $p = $p.Substring(0, $p.Length - 1)
    }

    return $p
}

function Get-RegNativePath {
    param([string]$PsPath)

    if ($PsPath -like 'HKLM:\*') { return ('HKLM\' + $PsPath.Substring(6)) }
    if ($PsPath -like 'HKCU:\*') { return ('HKCU\' + $PsPath.Substring(6)) }
    if ($PsPath -like 'HKCR:\*') { return ('HKCR\' + $PsPath.Substring(6)) }
    if ($PsPath -like 'HKU:\*')  { return ('HKU\'  + $PsPath.Substring(5)) }

    return $PsPath
}

function Get-RegSnapshot {
    param([string]$Path)

    $values = [ordered]@{}
    $exists = $false
    $err = $null
    $raw = $null

    try {
        if (Test-Path -LiteralPath $Path) {
            $exists = $true
            $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
            foreach ($p in $item.PSObject.Properties) {
                if ($p.Name -notmatch '^PS') {
                    $values[$p.Name] = $p.Value
                }
            }
        }
    }
    catch {
        $err = $_.Exception.Message
    }

    try {
        $native = Get-RegNativePath -PsPath $Path
        $raw = (& reg.exe query $native /s 2>&1 | Out-String).Trim()
    }
    catch {
        if (-not $err) {
            $err = $_.Exception.Message
        }
    }

    return [pscustomobject]@{
        Path   = $Path
        Exists = $exists
        Values = $values
        Raw    = $raw
        Error  = $err
    }
}

function Get-CimSafe {
    param([string]$ClassName)

    try {
        return @(Get-CimInstance -ClassName $ClassName -ErrorAction Stop)
    }
    catch {
        try {
            return @(Get-WmiObject -Class $ClassName -ErrorAction Stop)
        }
        catch {
            return @()
        }
    }
}

function Get-ShareInventory {
    return @(Get-CimSafe -ClassName Win32_Share | Select-Object Name, Path, Description, Type)
}

function Resolve-ShareForLocalPath {
    param(
        [string]$LocalPath,
        [array]$Shares,
        [string]$ComputerName
    )

    $normalized = Normalize-PathString $LocalPath
    if (-not $normalized) { return $null }

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
        if (-not $sharePath) { continue }
        if ($normalized.Length -lt $sharePath.Length) { continue }

        $starts = $normalized.StartsWith($sharePath, [System.StringComparison]::OrdinalIgnoreCase)
        if (-not $starts) { continue }

        $boundaryOk = $false
        if ($normalized.Length -eq $sharePath.Length) {
            $boundaryOk = $true
        }
        elseif ($normalized.Length -gt $sharePath.Length) {
            if ($normalized.Substring($sharePath.Length, 1) -eq '\') {
                $boundaryOk = $true
            }
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
    if ([string]::IsNullOrWhiteSpace($suffix)) {
        $unc = "\\$ComputerName\$($best.ShareName)"
    }
    else {
        $unc = "\\$ComputerName\$($best.ShareName)\$suffix"
    }

    return [pscustomobject]@{
        LocalPath = $normalized
        ShareName = $best.ShareName
        SharePath = $best.SharePath
        UncPath   = $unc
    }
}

function Get-ShareAccessSafe {
    param([string]$ShareName)

    if ([string]::IsNullOrWhiteSpace($ShareName)) { return $null }

    if (Get-Command Get-SmbShareAccess -ErrorAction SilentlyContinue) {
        try {
            return @(Get-SmbShareAccess -Name $ShareName -ErrorAction Stop |
                Select-Object Name, AccountName, AccessControlType, AccessRight)
        }
        catch {
            return $_.Exception.Message
        }
    }

    try {
        return (& net share $ShareName 2>&1 | Out-String).Trim()
    }
    catch {
        return 'Get-SmbShareAccess e net share indisponíveis.'
    }
}

function Get-CurrentNetworkConnections {
    return @(Get-CimSafe -ClassName Win32_NetworkConnection |
        Select-Object LocalName, RemoteName, UserName, ConnectionState, Status)
}

function Get-UserEnv {
    $vars = [ordered]@{}
    Get-ChildItem Env: | Sort-Object Name | ForEach-Object {
        $vars[$_.Name] = $_.Value
    }
    return $vars
}

function Get-PathInspection {
    param(
        [string]$Path,
        [switch]$IncludeAcl,
        [switch]$WriteProbe
    )

    $result = [ordered]@{
        Path       = $Path
        Exists     = $false
        IsUnc      = $false
        Kind       = $null
        Error      = $null
        Acl        = $null
        WriteProbe = $null
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        return [pscustomobject]$result
    }

    $result.IsUnc = ($Path -like '\\*')

    try {
        $item = Get-Item -LiteralPath $Path -Force -ErrorAction Stop
        $result.Exists = $true
        $result.Kind = if ($item.PSIsContainer) { 'Directory' } else { 'File' }

        if ($IncludeAcl) {
            try {
                $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
                $accessLines = @()
                foreach ($a in $acl.Access) {
                    $accessLines += ('{0} | {1} | {2}' -f $a.IdentityReference, $a.FileSystemRights, $a.AccessControlType)
                }

                $result.Acl = [pscustomobject]@{
                    Owner  = $acl.Owner
                    Access = $accessLines
                }
            }
            catch {
                $result.Acl = $_.Exception.Message
            }
        }

        if ($WriteProbe) {
            $probe = [ordered]@{
                Attempted = $true
                CanWrite  = $false
                ProbeFile = $null
                Error     = $null
            }

            try {
                $targetDir = if ($item.PSIsContainer) { $item.FullName } else { $item.DirectoryName }
                $probeFileName = [System.IO.Path]::GetRandomFileName()
                $probeFile = Join-Path $targetDir $probeFileName
                "probe $(Get-Date -Format o)" | Set-Content -LiteralPath $probeFile -Encoding ASCII -ErrorAction Stop
                Remove-Item -LiteralPath $probeFile -Force -ErrorAction Stop
                $probe.CanWrite  = $true
                $probe.ProbeFile = $probeFile
            }
            catch {
                $probe.Error = $_.Exception.Message
            }

            $result.WriteProbe = [pscustomobject]$probe
        }
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    return [pscustomobject]$result
}

function Get-SmbConfigSafe {
    $o = [ordered]@{
        SmbClient = $null
        SmbServer = $null
    }

    if (Get-Command Get-SmbClientConfiguration -ErrorAction SilentlyContinue) {
        try {
            $o.SmbClient = Get-SmbClientConfiguration | Select-Object `
                EnableSecuritySignature,
                RequireSecuritySignature,
                EnableInsecureGuestLogons,
                DirectoryCacheLifetime,
                FileInfoCacheLifetime,
                FileNotFoundCacheLifetime
        }
        catch {
            $o.SmbClient = $_.Exception.Message
        }
    }
    else {
        $o.SmbClient = 'Get-SmbClientConfiguration indisponível.'
    }

    if (Get-Command Get-SmbServerConfiguration -ErrorAction SilentlyContinue) {
        try {
            $o.SmbServer = Get-SmbServerConfiguration | Select-Object `
                EnableSMB1Protocol,
                EnableSMB2Protocol,
                RequireSecuritySignature,
                EnableSecuritySignature,
                EncryptData,
                RejectUnencryptedAccess
        }
        catch {
            $o.SmbServer = $_.Exception.Message
        }
    }
    else {
        $o.SmbServer = 'Get-SmbServerConfiguration indisponível.'
    }

    return [pscustomobject]$o
}

function Get-ProcessInventory {
    $out = @()

    try {
        $procs = Get-Process -ErrorAction Stop | Where-Object {
            $_.ProcessName -match 'ecg|heart|bde|idapi|paradox'
        }

        foreach ($p in $procs) {
            $path = $null
            $start = $null

            try { $path = $p.Path } catch {}
            try { $start = $p.StartTime } catch {}

            $out += [pscustomobject]@{
                ProcessName = $p.ProcessName
                Id          = $p.Id
                Path        = $path
                StartTime   = $start
            }
        }
    }
    catch {}

    return $out
}

function Add-UniquePath {
    param(
        [System.Collections.ArrayList]$List,
        [string]$Value
    )

    $p = Normalize-PathString $Value
    if ([string]::IsNullOrWhiteSpace($p)) { return }

    if (-not ($List -contains $p)) {
        [void]$List.Add($p)
    }
}

function Get-RegInterestingValues {
    param([array]$Snapshots)

    $items = @()
    $patterns = @(
        '^NET DIR$',
        '^LOCAL SHARE$',
        '^PRIVATE DIR$',
        '^LANGDRIVER$',
        '^VERSION$',
        'DATABASE',
        'CAMINHO',
        'NETDIR',
        'PRIVATE'
    )

    foreach ($snap in $Snapshots) {
        if (-not $snap.Exists) { continue }
        if ($null -eq $snap.Values) { continue }

        foreach ($key in $snap.Values.Keys) {
            $name = [string]$key
            $value = $snap.Values[$key]

            $match = $false
            foreach ($pat in $patterns) {
                if ($name -match $pat) {
                    $match = $true
                    break
                }
            }

            if ($match) {
                $items += [pscustomobject]@{
                    RegistryPath = $snap.Path
                    Name         = $name
                    Value        = $value
                }
            }
        }
    }

    return $items
}

function Get-HeartWareDbCandidates {
    param([array]$Snapshots)

    $items = @()

    foreach ($snap in $Snapshots) {
        if (-not $snap.Exists) { continue }
        if ($null -eq $snap.Values) { continue }

        foreach ($key in $snap.Values.Keys) {
            $name = [string]$key
            $value = $snap.Values[$key]
            $text = '{0} {1}' -f $name, $value

            if ($text -match 'database|caminho|hw\\|ecg') {
                $items += [pscustomobject]@{
                    RegistryPath = $snap.Path
                    Name         = $name
                    Value        = $value
                }
            }
        }
    }

    return $items
}

try {
    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
    }

    $now = Get-Date
    $stamp = $now.ToString('yyyyMMdd_HHmmss')
    $hostName = $env:COMPUTERNAME
    $userIdentity = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    $userLeaf = ($userIdentity -split '\\')[-1]
    $isAdmin = Test-IsAdmin

    $shares = @(Get-ShareInventory)
    $networkConns = @(Get-CurrentNetworkConnections)
    $userEnv = Get-UserEnv
    $smbConfig = Get-SmbConfigSafe
    $procInfo = @(Get-ProcessInventory)

    $envUserHwDb    = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::User)
    $envProcessHwDb = $env:HW_CAMINHO_DB
    $envMachineHwDb = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::Machine)

    $regPaths = @(
        'HKCU:\Environment',
        'HKCU:\Software\HeartWare\ECGV6',
        'HKCU:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\SOFTWARE\HeartWare\ECGV6',
        'HKLM:\SOFTWARE\WOW6432Node\HeartWare\ECGV6',
        'HKLM:\SOFTWARE\Borland\Database Engine',
        'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine',
        'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKLM:\SOFTWARE\Borland\Database Engine\Settings\DRIVERS\PARADOX\INIT',
        'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\DRIVERS\PARADOX\INIT',
        'HKCU:\Software\Borland\Database Engine',
        'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKCU:\Software\Borland\Database Engine\Settings\DRIVERS\PARADOX\INIT'
    )

    $regSnapshots = @()
    foreach ($rp in $regPaths) {
        $regSnapshots += Get-RegSnapshot -Path $rp
    }

    $interestingReg = @(Get-RegInterestingValues -Snapshots $regSnapshots)
    $heartWareCandidates = @(Get-HeartWareDbCandidates -Snapshots $regSnapshots)

    $hwDbChosen = $null
    if (-not [string]::IsNullOrWhiteSpace($envProcessHwDb)) {
        $hwDbChosen = $envProcessHwDb
    }
    elseif (-not [string]::IsNullOrWhiteSpace($envUserHwDb)) {
        $hwDbChosen = $envUserHwDb
    }
    elseif (-not [string]::IsNullOrWhiteSpace($envMachineHwDb)) {
        $hwDbChosen = $envMachineHwDb
    }

    $hwLocalToUnc = Resolve-ShareForLocalPath -LocalPath $hwDbChosen -Shares $shares -ComputerName $hostName

    $shareAccess = $null
    if ($hwLocalToUnc -and $hwLocalToUnc.ShareName) {
        $shareAccess = Get-ShareAccessSafe -ShareName $hwLocalToUnc.ShareName
    }

    $pathCandidates = New-Object System.Collections.ArrayList

    Add-UniquePath -List $pathCandidates -Value $hwDbChosen
    Add-UniquePath -List $pathCandidates -Value 'C:\HW'
    Add-UniquePath -List $pathCandidates -Value 'C:\HW\Database'
    Add-UniquePath -List $pathCandidates -Value 'C:\HW\NetDir'
    Add-UniquePath -List $pathCandidates -Value 'C:\HW\Private'

    # Filtro: adiciona apenas valores que parecem caminhos (local ou UNC)
    foreach ($i in $interestingReg) {
        if ($i.Value -is [string] -and ($i.Value -match '^[a-zA-Z]:\\|^\\\\')) {
            Add-UniquePath -List $pathCandidates -Value $i.Value
        }
    }

    foreach ($i in $heartWareCandidates) {
        if ($i.Value -is [string] -and ($i.Value -match '^[a-zA-Z]:\\|^\\\\')) {
            Add-UniquePath -List $pathCandidates -Value $i.Value
        }
    }

    if ($hwLocalToUnc -and $hwLocalToUnc.UncPath) {
        Add-UniquePath -List $pathCandidates -Value $hwLocalToUnc.UncPath
    }

    $shareHW = $shares | Where-Object { $_.Name -ieq 'HW' } | Select-Object -First 1
    if ($shareHW) {
        Add-UniquePath -List $pathCandidates -Value "\\$hostName\HW"
        Add-UniquePath -List $pathCandidates -Value "\\$hostName\HW\Database"
        Add-UniquePath -List $pathCandidates -Value "\\$hostName\HW\NetDir"
    }

    $pathInspection = @()
    foreach ($p in $pathCandidates) {
        $pathInspection += Get-PathInspection -Path ([string]$p) -IncludeAcl:$IncludeAcl -WriteProbe:$WriteProbe
    }

    $os = @(Get-CimSafe -ClassName Win32_OperatingSystem |
        Select-Object Caption, Version, BuildNumber, OSArchitecture, CSName)

    $cs = @(Get-CimSafe -ClassName Win32_ComputerSystem |
        Select-Object Name, Domain, Manufacturer, Model, UserName)

    $ip = @(Get-CimSafe -ClassName Win32_NetworkAdapterConfiguration |
        Where-Object { $_.IPEnabled -eq $true } |
        Select-Object Description, IPAddress, DefaultIPGateway, MACAddress)

    # --- Geração de alertas ---
    $alerts = @()
    if ([string]::IsNullOrWhiteSpace($hwDbChosen)) {
        $alerts += "AVISO: Nenhuma definição da variável HW_CAMINHO_DB encontrada (processo, usuário ou máquina)."
    }
    if ($hwLocalToUnc -and -not $hwLocalToUnc.ShareName) {
        $alerts += "AVISO: HW_CAMINHO_DB não pôde ser resolvido para um caminho UNC (nenhum compartilhamento local correspondente)."
    }
    $missingPaths = $pathInspection | Where-Object { -not $_.Exists } | ForEach-Object { $_.Path }
    if ($missingPaths) {
        $alerts += "AVISO: Os seguintes caminhos não existem: $($missingPaths -join ', ')"
    }
    $writeFailed = $pathInspection | Where-Object { $_.Exists -and $WriteProbe -and $_.WriteProbe -and -not $_.WriteProbe.CanWrite } | ForEach-Object { $_.Path }
    if ($writeFailed) {
        $alerts += "AVISO: Sem permissão de escrita nos seguintes caminhos: $($writeFailed -join ', ')"
    }

    $summary = [ordered]@{
        Timestamp                 = $now.ToString('yyyy-MM-dd HH:mm:ss')
        Host                      = $hostName
        User                      = $userIdentity
        IsAdmin                   = $isAdmin
        HW_CAMINHO_DB_ProcessEnv  = $envProcessHwDb
        HW_CAMINHO_DB_UserEnv     = $envUserHwDb
        HW_CAMINHO_DB_MachineEnv  = $envMachineHwDb
        HW_CAMINHO_DB_Effective   = $hwDbChosen
        HW_CAMINHO_DB_DerivedUNC  = $(if ($hwLocalToUnc) { $hwLocalToUnc.UncPath } else { $null })
        LocalHWShareMatched       = $(if ($hwLocalToUnc) { $hwLocalToUnc.ShareName } else { $null })
        InterestingRegistryValues = $interestingReg
        HeartWareDbCandidates     = $heartWareCandidates
        PathsChecked              = @($pathCandidates)
        ECGProcesses              = $procInfo
        Alerts                    = $alerts
    }

    $report = [ordered]@{
        Meta = [ordered]@{
            GeneratedAt = $now.ToString('o')
            Host        = $hostName
            User        = $userIdentity
            UserOnly    = $userLeaf
            Domain      = $env:USERDOMAIN
            IsAdmin     = $isAdmin
            OutDir      = $OutDir
            IncludeAcl  = [bool]$IncludeAcl
            WriteProbe  = [bool]$WriteProbe
            ScriptName  = 'Get-ECGv6-BDE-State.ps1'
        }

        Summary            = $summary
        ComputerSystem     = $cs
        OperatingSystem    = $os
        IPConfiguration    = $ip
        UserEnvironment    = $userEnv
        Shares             = $shares
        ShareAccess        = $shareAccess
        NetworkConnections = $networkConns
        SmbConfiguration   = $smbConfig
        RelevantProcesses  = $procInfo
        PathInspection     = $pathInspection
        RegistrySnapshots  = $regSnapshots
    }

    $baseName = 'ECG_State_{0}_{1}_{2}' -f $hostName, $userLeaf, $stamp
    $jsonFile = Join-Path $OutDir ($baseName + '.json')
    $txtFile  = Join-Path $OutDir ($baseName + '.txt')

    $report | ConvertTo-Json -Depth 12 | Set-Content -LiteralPath $jsonFile -Encoding UTF8

    # --- Construção do TXT (compatível com PowerShell 5.1) ---
    $txt = @()
    $txt += '=' * 60
    $txt += 'ECG/BDE STATE REPORT'
    $txt += '=' * 60
    $txt += ('GeneratedAt : {0}' -f $report.Meta.GeneratedAt)
    $txt += ('Host        : {0}' -f $report.Meta.Host)
    $txt += ('User        : {0}' -f $report.Meta.User)
    $txt += ('IsAdmin     : {0}' -f $report.Meta.IsAdmin)
    $txt += ''

    if ($alerts.Count -gt 0) {
        $txt += 'ALERTAS'
        $txt += '-------'
        $txt += $alerts
        $txt += ''
    }

    $txt += 'HW_CAMINHO_DB'
    $txt += '------------'
    $txt += ('Process Env : {0}' -f $(if ($summary.HW_CAMINHO_DB_ProcessEnv) { $summary.HW_CAMINHO_DB_ProcessEnv } else { '<não definido>' }))
    $txt += ('User Env    : {0}' -f $(if ($summary.HW_CAMINHO_DB_UserEnv) { $summary.HW_CAMINHO_DB_UserEnv } else { '<não definido>' }))
    $txt += ('Machine Env : {0}' -f $(if ($summary.HW_CAMINHO_DB_MachineEnv) { $summary.HW_CAMINHO_DB_MachineEnv } else { '<não definido>' }))
    $txt += ('Effective   : {0}' -f $(if ($summary.HW_CAMINHO_DB_Effective) { $summary.HW_CAMINHO_DB_Effective } else { '<não definido>' }))
    $txt += ('Derived UNC : {0}' -f $(if ($summary.HW_CAMINHO_DB_DerivedUNC) { $summary.HW_CAMINHO_DB_DerivedUNC } else { '<não resolvido>' }))
    $txt += ('Local Share : {0}' -f $(if ($summary.LocalHWShareMatched) { $summary.LocalHWShareMatched } else { '<nenhum>' }))
    $txt += ''

    $txt += 'INTERESTING REGISTRY VALUES'
    $txt += '--------------------------'
    if ($interestingReg.Count -gt 0) {
        foreach ($i in $interestingReg) {
            $txt += ('[{0}] {1} = {2}' -f $i.RegistryPath, $i.Name, $i.Value)
        }
    }
    else {
        $txt += '<nenhum valor relevante encontrado>'
    }
    $txt += ''

    $txt += 'HEARTWARE DB CANDIDATES'
    $txt += '-----------------------'
    if ($heartWareCandidates.Count -gt 0) {
        foreach ($i in $heartWareCandidates) {
            $txt += ('[{0}] {1} = {2}' -f $i.RegistryPath, $i.Name, $i.Value)
        }
    }
    else {
        $txt += '<nenhum candidato encontrado>'
    }
    $txt += ''

    $txt += 'PROCESSOS RELEVANTES'
    $txt += '--------------------'
    if ($procInfo.Count -gt 0) {
        foreach ($p in $procInfo) {
            $pathDisplay = if ($p.Path) { $p.Path } else { '<desconhecido>' }
            $startDisplay = if ($p.StartTime) { $p.StartTime } else { '<não disponível>' }
            $txt += ('{0} (PID {1}) - Path: {2} - Iniciado: {3}' -f $p.ProcessName, $p.Id, $pathDisplay, $startDisplay)
        }
    }
    else {
        $txt += '<nenhum processo ECG/Heart/BDE encontrado>'
    }
    $txt += ''

    $txt += 'CONFIGURAÇÃO SMB'
    $txt += '----------------'
    $txt += 'Cliente:'
    if ($smbConfig.SmbClient -is [string]) {
        $txt += "  $($smbConfig.SmbClient)"
    }
    else {
        $smbConfig.SmbClient.PSObject.Properties | ForEach-Object {
            $txt += "  $($_.Name) = $($_.Value)"
        }
    }
    $txt += 'Servidor:'
    if ($smbConfig.SmbServer -is [string]) {
        $txt += "  $($smbConfig.SmbServer)"
    }
    else {
        $smbConfig.SmbServer.PSObject.Properties | ForEach-Object {
            $txt += "  $($_.Name) = $($_.Value)"
        }
    }
    $txt += ''

    $txt += 'INSPEÇÃO DE CAMINHOS'
    $txt += '--------------------'
    foreach ($p in $pathInspection) {
        $status = if ($p.Exists) { "Existe ($($p.Kind))" } else { "NÃO EXISTE" }
        $line = ("{0} | {1}" -f $p.Path, $status)
        if ($p.Error) { $line += " | Erro: $($p.Error)" }
        if ($WriteProbe -and $p.Exists -and $p.WriteProbe -and -not $p.WriteProbe.CanWrite) {
            $line += " | Sem permissão de escrita"
        }
        $txt += $line
    }
    $txt += ''

    $txt += 'COMPARTILHAMENTOS'
    $txt += '-----------------'
    if ($shares.Count -gt 0) {
        foreach ($s in $shares) {
            $txt += ('{0} | {1} | {2}' -f $s.Name, $s.Path, $s.Description)
        }
    }
    else {
        $txt += '<nenhum compartilhamento listado>'
    }
    $txt += ''

    $txt += 'CONEXÕES DE REDE'
    $txt += '----------------'
    if ($networkConns.Count -gt 0) {
        foreach ($n in $networkConns) {
            $txt += ('{0} -> {1} | Usuário={2} | Estado={3} | Status={4}' -f $n.LocalName, $n.RemoteName, $n.UserName, $n.ConnectionState, $n.Status)
        }
    }
    else {
        $txt += '<nenhum mapeamento listado>'
    }
    $txt += ''

    $txt += 'ARQUIVOS GERADOS'
    $txt += '----------------'
    $txt += $jsonFile
    $txt += $txtFile
    $txt += '=' * 60

    $txt -join [Environment]::NewLine | Set-Content -LiteralPath $txtFile -Encoding UTF8

    Write-Host ''
    Write-Host 'OK - coleta concluída.' -ForegroundColor Green
    Write-Host ('JSON: {0}' -f $jsonFile)
    Write-Host ('TXT : {0}' -f $txtFile)
    Write-Host ''
}
catch {
    Write-Host ''
    Write-Host ('ERRO na linha {0}: {1}' -f $_.InvocationInfo.ScriptLineNumber, $_.Exception.Message) -ForegroundColor Red
    throw
}