<#
.SYNOPSIS
    Ferramenta autônoma de diagnóstico e correção controlada para ECGv6/BDE em cenário legado,
    incluindo laboratório com VM Windows XP SP3 x86 como host de compatibilidade do DBE.

.DESCRIPTION
    Objetivos da ferramenta:
      - descobrir e consolidar caminho efetivo do banco (HW_CAMINHO_DB)
      - descobrir e consolidar NETDIR do BDE
      - validar acesso UNC direto, consistência de configuração, permissão de escrita no NetDir
      - inspecionar share local (quando executada no host do banco)
      - observar estabilidade temporal de DB/NetDir e contagem de arquivos de lock
      - aplicar correções idempotentes e seguras quando o alvo esperado for determinístico
      - gerar laudo TXT, JSON e HTML por rodada
      - CORREÇÃO AUTOMÁTICA: criar chave HKCU do BDE e ajustar IDAPI32.CFG

    Compatibilidade alvo:
      - Windows PowerShell 2.0+
      - Windows XP SP3 x86 (modo reduzido, sem dependência de CIM/ConvertTo-Json)
      - Windows 7/10/11 e Windows Server (modo pleno)

.NOTES
    - Não usa Invoke-Expression.
    - Não usa ConvertTo-Json nativo.
    - Não depende de módulos modernos.
    - Prioriza UNC direto e evita unidade mapeada como caminho canônico.
#>

[CmdletBinding(SupportsShouldProcess=$true, ConfirmImpact='Medium')]
param(
    [ValidateSet('Prepare','Audit','Auto','Fix','Compare','Rollback')]
    [string]$Mode = 'Auto',

    [string]$ExpectedDbPath = '',
    [string]$ExpectedNetDir = '',
    [string]$ExpectedExePath = '',
    [string]$ProfilePath = '',
    [string]$OutDir = 'C:\ECG\FieldKit',
    [string]$SymptomText = '',
    [string]$CompareLeftReport = '',
    [string]$CompareRightReport = '',
    [string]$RollbackFile = '',

    [ValidateRange(0,120)]
    [int]$MonitorMinutes = 3,

    [ValidateRange(5,300)]
    [int]$SampleIntervalSeconds = 15,

    [switch]$WriteProbe,
    [switch]$IncludeAcl,
    [switch]$CreateMissingDirs,
    [switch]$SetMachineHwPath,
    [switch]$OpenReport,
    [switch]$Elevated
)

$ErrorActionPreference = 'Stop'
$script:ToolName = 'ECGv6 FieldKit'
$script:ToolVersion = 'FINAL-2026-04-04-r2'
$script:RunStartedAt = Get-Date
$script:HostName = $env:COMPUTERNAME
$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:LogLines = New-Object System.Collections.ArrayList
$script:RecommendationLines = New-Object System.Collections.ArrayList
$script:AppliedChanges = New-Object System.Collections.ArrayList
$script:PendingChanges = New-Object System.Collections.ArrayList
$script:Warnings = New-Object System.Collections.ArrayList
$script:ErrorsFound = New-Object System.Collections.ArrayList

function Add-ListItem {
    param(
        [Parameter(Mandatory=$true)][System.Collections.IList]$List,
        [Parameter(Mandatory=$true)]$Value
    )
    if ($null -ne $Value) {
        [void]$List.Add($Value)
    }
}

function New-Map {
    return New-Object 'System.Collections.Hashtable'
}

function New-List {
    return New-Object 'System.Collections.ArrayList'
}

function Log {
    param(
        [string]$Message,
        [string]$Level = 'INFO'
    )
    $line = ('{0} [{1}] {2}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message)
    [void]$script:LogLines.Add($line)
    Write-Host $line
}

function Warn {
    param([string]$Message)
    Add-ListItem -List $script:Warnings -Value $Message
    Log -Message $Message -Level 'WARN'
}

function Fail-Note {
    param([string]$Message)
    Add-ListItem -List $script:ErrorsFound -Value $Message
    Log -Message $Message -Level 'ERROR'
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

function Test-IsUncPath {
    param([string]$Path)
    $p = Normalize-PathString $Path
    if ($null -eq $p) { return $false }
    return $p.StartsWith('\\')
}

function Test-IsDrivePath {
    param([string]$Path)
    $p = Normalize-PathString $Path
    if ($null -eq $p) { return $false }
    return ($p -match '^[A-Za-z]:\\')
}

function Convert-DrivePathToUnc {
    param(
        [string]$Path,
        [array]$NetworkConnections
    )
    $normalized = Normalize-PathString $Path
    if ($null -eq $normalized) { return $null }
    if (-not (Test-IsDrivePath $normalized)) { return $normalized }

    $drive = $normalized.Substring(0,2).ToUpperInvariant()
    $suffix = ''
    if ($normalized.Length -gt 3) {
        $suffix = $normalized.Substring(3)
    }

    foreach ($conn in @($NetworkConnections)) {
        $localName = $null
        $remoteName = $null
        try { $localName = $conn.LocalName } catch {}
        try { $remoteName = $conn.RemoteName } catch {}
        if ([string]::IsNullOrWhiteSpace($localName) -or [string]::IsNullOrWhiteSpace($remoteName)) { continue }
        if ($localName.ToUpperInvariant() -eq $drive) {
            if ([string]::IsNullOrWhiteSpace($suffix)) {
                return (Normalize-PathString $remoteName)
            }
            return (Normalize-PathString ($remoteName.TrimEnd('\\') + '\\' + $suffix))
        }
    }

    return $normalized
}

function Resolve-ShareForLocalPath {
    param(
        [string]$LocalPath,
        [array]$Shares,
        [string]$ComputerName
    )

    $normalized = Normalize-PathString $LocalPath
    if ($null -eq $normalized) { return $null }

    if (Test-IsUncPath $normalized) {
        $m = New-Map
        $m.LocalPath = $normalized
        $m.ShareName = $null
        $m.SharePath = $null
        $m.UncPath = $normalized
        return $m
    }

    $bestShare = $null
    $bestSharePath = $null
    foreach ($share in @($Shares)) {
        $shareName = $null
        $sharePath = $null
        try { $shareName = $share.Name } catch {}
        try { $sharePath = $share.Path } catch {}
        $sharePath = Normalize-PathString $sharePath
        if ([string]::IsNullOrWhiteSpace($shareName) -or [string]::IsNullOrWhiteSpace($sharePath)) { continue }
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

        if (($null -eq $bestSharePath) -or ($sharePath.Length -gt $bestSharePath.Length)) {
            $bestShare = $shareName
            $bestSharePath = $sharePath
        }
    }

    if ($null -eq $bestShare) { return $null }

    $suffix = ''
    if ($normalized.Length -gt $bestSharePath.Length) {
        $suffix = $normalized.Substring($bestSharePath.Length).TrimStart('\\')
    }
    $unc = ('\\{0}\{1}' -f $ComputerName, $bestShare)
    if (-not [string]::IsNullOrWhiteSpace($suffix)) {
        $unc = $unc + '\\' + $suffix
    }

    $result = New-Map
    $result.LocalPath = $normalized
    $result.ShareName = $bestShare
    $result.SharePath = $bestSharePath
    $result.UncPath = $unc
    return $result
}

function Get-WmiSafe {
    param([string]$ClassName)
    try {
        return @(Get-WmiObject -Class $ClassName -ErrorAction Stop)
    }
    catch {
        return @()
    }
}

function Get-RegistryValueSafe {
    param(
        [string]$Path,
        [string]$Name
    )
    try {
        $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
        return $item.$Name
    }
    catch {
        return $null
    }
}

function Get-RegistrySnapshot {
    param([string]$Path)
    $obj = New-Map
    $obj.Path = $Path
    $obj.Exists = $false
    $obj.Values = New-Map
    $obj.Error = $null
    try {
        if (Test-Path -LiteralPath $Path) {
            $obj.Exists = $true
            $item = Get-ItemProperty -LiteralPath $Path -ErrorAction Stop
            foreach ($prop in $item.PSObject.Properties) {
                if ($prop.Name -notmatch '^PS') {
                    $obj.Values[$prop.Name] = $prop.Value
                }
            }
        }
    }
    catch {
        $obj.Error = $_.Exception.Message
    }
    return $obj
}

function Get-UserAndMachineEnvState {
    $obj = New-Map
    $obj.Process = $env:HW_CAMINHO_DB
    $obj.User = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', 'User')
    $obj.Machine = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', 'Machine')
    return $obj
}

function Get-InterestingDbCandidates {
    param([array]$Snapshots)
    $result = New-List
    $seen = New-Map
    foreach ($snap in @($Snapshots)) {
        if (-not $snap.Exists) { continue }
        foreach ($entry in $snap.Values.GetEnumerator()) {
            $name = [string]$entry.Key
            $value = $entry.Value
            if ($value -isnot [string]) { continue }
            if ($value -match '^[A-Za-z]:\\|^\\\\') {
                $id = ($snap.Path + '|' + $name + '|' + $value)
                if (-not $seen.ContainsKey($id)) {
                    $seen[$id] = $true
                    $row = New-Map
                    $row.RegistryPath = $snap.Path
                    $row.Name = $name
                    $row.Value = $value
                    [void]$result.Add($row)
                }
            }
        }
    }
    return $result
}

function Get-BdeNetDirState {
    $roots = @(
        'HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT',
        'HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT'
    )

    $result = New-List
    foreach ($root in $roots) {
        $row = New-Map
        $row.RegistryPath = $root
        $row.Exists = [bool](Test-Path -LiteralPath $root)
        $row.CurrentValue = Get-RegistryValueSafe -Path $root -Name 'NETDIR'
        [void]$result.Add($row)
    }
    return $result
}

function Get-HeartWareRegistryState {
    $paths = @(
        'HKCU:\Software\HeartWare\ECGV6',
        'HKCU:\Software\HeartWare\ECGV6\Geral',
        'HKLM:\SOFTWARE\HeartWare\ECGV6',
        'HKLM:\SOFTWARE\WOW6432Node\HeartWare\ECGV6'
    )
    $result = New-List
    foreach ($path in $paths) {
        [void]$result.Add((Get-RegistrySnapshot -Path $path))
    }
    return $result
}

function Get-LocalShares {
    $shares = Get-WmiSafe -ClassName 'Win32_Share'
    $result = New-List
    foreach ($s in @($shares)) {
        $row = New-Map
        $row.Name = $s.Name
        $row.Path = $s.Path
        $row.Description = $s.Description
        $row.Type = $s.Type
        [void]$result.Add($row)
    }
    return $result
}

function Get-NetworkConnections {
    $items = Get-WmiSafe -ClassName 'Win32_NetworkConnection'
    $result = New-List
    foreach ($n in @($items)) {
        $row = New-Map
        $row.LocalName = $n.LocalName
        $row.RemoteName = $n.RemoteName
        $row.UserName = $n.UserName
        $row.ConnectionState = $n.ConnectionState
        $row.Status = $n.Status
        [void]$result.Add($row)
    }
    return $result
}

function Get-OperatingSystemState {
    $os = @()
    try { $os = @(Get-WmiObject Win32_OperatingSystem -ErrorAction Stop) } catch { $os = @() }
    if ($os.Count -gt 0) {
        $item = $os[0]
        $obj = New-Map
        $obj.Caption = $item.Caption
        $obj.Version = $item.Version
        $obj.BuildNumber = $item.BuildNumber
        $obj.OSArchitecture = $item.OSArchitecture
        $obj.CSName = $item.CSName
        return $obj
    }
    return $null
}

function Get-ComputerSystemState {
    $cs = @()
    try { $cs = @(Get-WmiObject Win32_ComputerSystem -ErrorAction Stop) } catch { $cs = @() }
    if ($cs.Count -gt 0) {
        $item = $cs[0]
        $obj = New-Map
        $obj.Name = $item.Name
        $obj.Domain = $item.Domain
        $obj.Manufacturer = $item.Manufacturer
        $obj.Model = $item.Model
        $obj.UserName = $item.UserName
        return $obj
    }
    return $null
}

function Get-CurrentIdentity {
    $obj = New-Map
    try {
        $obj.Name = [Security.Principal.WindowsIdentity]::GetCurrent().Name
    }
    catch {
        $obj.Name = $env:USERNAME
    }
    $obj.UserName = $env:USERNAME
    $obj.Domain = $env:USERDOMAIN
    return $obj
}

function Get-ProcessInventory {
    $result = New-List
    try {
        $processes = @(Get-Process -ErrorAction Stop)
    }
    catch {
        $processes = @()
    }

    foreach ($proc in @($processes)) {
        $name = $null
        try { $name = $proc.ProcessName } catch {}
        if ([string]::IsNullOrWhiteSpace($name)) { continue }
        if ($name -notmatch 'ECG|BDE|IDAPI|Heart') { continue }

        $row = New-Map
        $row.ProcessName = $name
        try { $row.Id = $proc.Id } catch {}
        try { $row.Path = $proc.Path } catch {}
        try { $row.StartTime = $proc.StartTime } catch {}
        [void]$result.Add($row)
    }
    return $result
}

function Get-RoleHint {
    param(
        $Profile,
        [array]$Shares,
        [string]$DesiredDbPath,
        [string]$ExpectedExePath,
        $OsState
    )

    $profileRole = Get-ProfileRole -Profile $Profile
    if (-not [string]::IsNullOrWhiteSpace($profileRole)) {
        return $profileRole
    }

    $hasHwShare = $false
    foreach ($share in @($Shares)) {
        if ([string]$share.Name -ieq 'HW') {
            $hasHwShare = $true
            break
        }
    }

    if ($hasHwShare) {
        if ($OsState -and ([string]$OsState.Caption -match 'XP')) {
            return 'Storage XP/compatibilidade'
        }
        return 'Servidor de banco/compatibilidade'
    }

    if (-not [string]::IsNullOrWhiteSpace($ExpectedExePath) -and (Test-Path -LiteralPath $ExpectedExePath)) {
        return 'Estação executante'
    }

    return 'Estação visualizadora/cliente'
}

function Read-TextFileSafe {
    param([string]$Path)
    try {
        return [System.IO.File]::ReadAllText($Path)
    }
    catch {
        return $null
    }
}

function Get-IdapiCfgState {
    $paths = @(
        'C:\Program Files\Common Files\Borland Shared\BDE\IDAPI32.CFG',
        'C:\Program Files (x86)\Common Files\Borland Shared\BDE\IDAPI32.CFG',
        'C:\Arquivos de Programas\Arquivos Comuns\Borland Shared\BDE\IDAPI32.CFG'
    )

    $result = New-List
    foreach ($path in $paths) {
        $row = New-Map
        $row.Path = $path
        $row.Exists = [bool](Test-Path -LiteralPath $path)
        $row.ContainsNetDir = $false
        $row.Extract = $null
        if ($row.Exists) {
            $text = Read-TextFileSafe -Path $path
            if ($text) {
                $row.ContainsNetDir = [bool]($text -match 'NET\s*DIR|NETDIR')
                if ($text.Length -gt 1200) {
                    $row.Extract = $text.Substring(0, 1200)
                }
                else {
                    $row.Extract = $text
                }
            }
        }
        [void]$result.Add($row)
    }
    return $result
}

function Get-PathAclSafe {
    param([string]$Path)
    try {
        $acl = Get-Acl -LiteralPath $Path -ErrorAction Stop
        $result = New-List
        foreach ($rule in @($acl.Access)) {
            $row = New-Map
            $row.IdentityReference = [string]$rule.IdentityReference
            $row.FileSystemRights = [string]$rule.FileSystemRights
            $row.AccessControlType = [string]$rule.AccessControlType
            $row.IsInherited = [bool]$rule.IsInherited
            [void]$result.Add($row)
        }
        return $result
    }
    catch {
        return $null
    }
}

function Invoke-WriteProbe {
    param([string]$DirectoryPath)
    $probe = New-Map
    $probe.Attempted = $false
    $probe.CanWrite = $false
    $probe.Error = $null
    $probe.ProbeFile = $null

    if ([string]::IsNullOrWhiteSpace($DirectoryPath)) {
        return $probe
    }

    if (-not (Test-Path -LiteralPath $DirectoryPath)) {
        $probe.Error = 'Diretório inexistente.'
        return $probe
    }

    $probe.Attempted = $true
    $probeFile = Join-Path $DirectoryPath ('ecgv6_probe_' + [Guid]::NewGuid().ToString('N') + '.tmp')
    $probe.ProbeFile = $probeFile
    try {
        [System.IO.File]::WriteAllText($probeFile, 'probe=' + (Get-Date).ToString('o'))
        $probe.CanWrite = $true
    }
    catch {
        $probe.Error = $_.Exception.Message
    }
    finally {
        try {
            if (Test-Path -LiteralPath $probeFile) {
                Remove-Item -LiteralPath $probeFile -Force -ErrorAction SilentlyContinue
            }
        }
        catch {}
    }
    return $probe
}

function Get-LockInventory {
    param([string]$NetDirPath)
    $result = New-Map
    $result.Path = $NetDirPath
    $result.Count = 0
    $result.Files = New-List
    if ([string]::IsNullOrWhiteSpace($NetDirPath)) {
        return $result
    }
    try {
        if (Test-Path -LiteralPath $NetDirPath) {
            $files = @(Get-ChildItem -LiteralPath $NetDirPath -Force -ErrorAction SilentlyContinue)
            foreach ($file in @($files)) {
                if ($file.Name -match 'PDOXUSRS|PARADOX|\.LCK$|\.NET$') {
                    $row = New-Map
                    $row.Name = $file.Name
                    $row.Length = $file.Length
                    $row.LastWriteTime = $file.LastWriteTime
                    [void]$result.Files.Add($row)
                }
            }
            $result.Count = $result.Files.Count
        }
    }
    catch {}
    return $result
}

function Inspect-Path {
    param(
        [string]$Path,
        [switch]$DoWriteProbe,
        [switch]$DoAcl
    )

    $obj = New-Map
    $normalized = Normalize-PathString $Path
    $obj.Path = $normalized
    $obj.Exists = $false
    $obj.PathType = 'UNKNOWN'
    $obj.ParentExists = $false
    $obj.IsUnc = $false
    $obj.IsDrivePath = $false
    $obj.ItemType = $null
    $obj.Error = $null
    $obj.WriteProbe = $null
    $obj.Acl = $null

    if ($null -eq $normalized) {
        $obj.Error = 'Caminho vazio.'
        return $obj
    }

    $obj.IsUnc = Test-IsUncPath $normalized
    $obj.IsDrivePath = Test-IsDrivePath $normalized
    if ($obj.IsUnc) { $obj.PathType = 'UNC' }
    elseif ($obj.IsDrivePath) { $obj.PathType = 'LOCAL_OR_MAPPED' }

    try {
        $parent = Split-Path -Path $normalized -Parent
        if (-not [string]::IsNullOrWhiteSpace($parent)) {
            $obj.ParentExists = [bool](Test-Path -LiteralPath $parent)
        }
    }
    catch {}

    try {
        if (Test-Path -LiteralPath $normalized) {
            $obj.Exists = $true
            $item = Get-Item -LiteralPath $normalized -ErrorAction Stop
            $obj.ItemType = $item.PSProvider.Name
            if ($item -is [System.IO.DirectoryInfo]) {
                $obj.ItemKind = 'Directory'
            }
            elseif ($item -is [System.IO.FileInfo]) {
                $obj.ItemKind = 'File'
            }
            else {
                $obj.ItemKind = 'Other'
            }
        }
    }
    catch {
        $obj.Error = $_.Exception.Message
    }

    if ($obj.Exists -and $DoWriteProbe -and $obj.ItemKind -eq 'Directory') {
        $obj.WriteProbe = Invoke-WriteProbe -DirectoryPath $normalized
    }

    if ($obj.Exists -and $DoAcl) {
        $obj.Acl = Get-PathAclSafe -Path $normalized
    }

    return $obj
}

function Get-ProfileValueMap {
    param([string]$Path)
    $map = New-Map
    if ([string]::IsNullOrWhiteSpace($Path)) { return $map }
    if (-not (Test-Path -LiteralPath $Path)) { return $map }

    $lines = Get-Content -LiteralPath $Path -ErrorAction Stop
    foreach ($line in @($lines)) {
        $trimmed = [string]$line
        if ([string]::IsNullOrWhiteSpace($trimmed)) { continue }
        $trimmed = $trimmed.Trim()
        if ($trimmed.StartsWith('#') -or $trimmed.StartsWith(';')) { continue }
        $idx = $trimmed.IndexOf('=')
        if ($idx -lt 1) { continue }
        $key = $trimmed.Substring(0, $idx).Trim()
        $value = $trimmed.Substring($idx + 1).Trim()
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $map[$key] = $value
        }
    }
    return $map
}

function Coalesce-NonEmpty {
    param([object[]]$Values)
    foreach ($value in @($Values)) {
        if ($value -is [string]) {
            if (-not [string]::IsNullOrWhiteSpace($value)) {
                return $value
            }
        }
        elseif ($null -ne $value) {
            return $value
        }
    }
    return $null
}


function Get-MapValue {
    param(
        $Map,
        [string]$Key,
        [object]$Default = $null
    )
    try {
        if ($null -ne $Map -and $Map.ContainsKey($Key)) {
            $value = $Map[$Key]
            if ($value -is [string]) {
                if (-not [string]::IsNullOrWhiteSpace($value)) { return $value }
            }
            elseif ($null -ne $value) {
                return $value
            }
        }
    }
    catch {}
    return $Default
}

function Get-ProfileRole {
    param($Profile)
    $raw = [string](Get-MapValue -Map $Profile -Key 'StationRole' -Default '')
    $role = $raw.Trim().ToUpperInvariant()
    switch ($role) {
        'EXECUTANTE' { return 'Estação executante' }
        'VIEWER' { return 'Estação visualizadora/cliente' }
        'HOST_XP' { return 'Storage XP/compatibilidade' }
        'XP_STORAGE' { return 'Storage XP/compatibilidade' }
        'AUTO' { return $null }
        'UNDEFINED' { return $null }
        default { return $null }
    }
}

function Get-ProfileAlias {
    param($Profile)
    return [string](Get-MapValue -Map $Profile -Key 'StationAlias' -Default '')
}

function Get-ProfileOutDir {
    param($Profile)
    return [string](Get-MapValue -Map $Profile -Key 'OutDir' -Default $OutDir)
}

function Get-BdeNetDirMap {
    param([array]$BdeNetDirState)
    $map = New-Map
    foreach ($row in @($BdeNetDirState)) {
        $key = [string]$row.RegistryPath
        $value = Normalize-PathString ([string]$row.CurrentValue)
        if (-not [string]::IsNullOrWhiteSpace($key)) {
            $map[$key] = $value
        }
    }
    return $map
}

function Get-HeartWareDbPath {
    param([array]$HeartWareSnapshots)
    foreach ($snap in @($HeartWareSnapshots)) {
        try {
            if ($snap.Exists -and $snap.Values.ContainsKey('Caminho Database')) {
                $v = Normalize-PathString ([string]$snap.Values['Caminho Database'])
                if (-not [string]::IsNullOrWhiteSpace($v)) { return $v }
            }
        }
        catch {}
    }
    return $null
}

function Get-IdapiNetDirHints {
    param([array]$IdapiCfgState)
    $list = New-List
    foreach ($cfg in @($IdapiCfgState)) {
        if ($cfg.Exists -and -not [string]::IsNullOrWhiteSpace([string]$cfg.Extract)) {
            $extract = [string]$cfg.Extract
            $normalized = $extract -replace [char]0, ' '
            if ($normalized -match 'NET\s*DIR') {
                [void]$list.Add($normalized)
            }
        }
    }
    return $list
}

function Read-JsonReportObject {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Caminho de relatório não informado.' }
    if (-not (Test-Path -LiteralPath $Path)) { throw ('Relatório não encontrado: ' + $Path) }
    $raw = [System.IO.File]::ReadAllText($Path)
    try {
        if (Get-Command ConvertFrom-Json -ErrorAction SilentlyContinue) {
            return ($raw | ConvertFrom-Json -ErrorAction Stop)
        }
    }
    catch {}
    try {
        Add-Type -AssemblyName System.Web.Extensions -ErrorAction SilentlyContinue | Out-Null
        $ser = New-Object System.Web.Script.Serialization.JavaScriptSerializer
        $ser.MaxJsonLength = 67108864
        return $ser.DeserializeObject($raw)
    }
    catch {
        throw ('Falha ao ler JSON do relatório: ' + $_.Exception.Message)
    }
}

function Get-ReportValue {
    param($Object, [string[]]$PathParts)
    $current = $Object
    foreach ($part in @($PathParts)) {
        if ($null -eq $current) { return $null }
        try {
            if ($current -is [System.Collections.IDictionary]) {
                if ($current.Contains($part)) { $current = $current[$part] }
                elseif ($current.ContainsKey($part)) { $current = $current[$part] }
                else { return $null }
            }
            else {
                $prop = $current.PSObject.Properties[$part]
                if ($null -eq $prop) { return $null }
                $current = $prop.Value
            }
        }
        catch { return $null }
    }
    return $current
}


function Resolve-CompareReportPaths {
    param(
        [string]$RootOutDir,
        [string]$LeftPath,
        [string]$RightPath
    )
    $result = New-Map
    $result.Left = $LeftPath
    $result.Right = $RightPath

    if (-not [string]::IsNullOrWhiteSpace([string]$result.Left) -and -not [string]::IsNullOrWhiteSpace([string]$result.Right)) {
        return $result
    }
    if ([string]::IsNullOrWhiteSpace([string]$RootOutDir)) {
        throw 'Compare requer OutDir válido ou os 2 relatórios explicitamente.'
    }
    if (-not (Test-Path -LiteralPath $RootOutDir)) {
        throw ('OutDir não encontrado para comparação: ' + $RootOutDir)
    }
    $files = @(Get-ChildItem -LiteralPath $RootOutDir -Recurse -Filter 'ECGv6_FieldKit_Report.json' -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending)
    if ($files.Count -lt 2) {
        throw ('Foram encontrados apenas ' + [string]$files.Count + ' relatório(s) JSON em ' + $RootOutDir + '. São necessários 2.')
    }
    $result.Left = $files[0].FullName
    $result.Right = $files[1].FullName
    return $result
}

function Prepare-ExpectedPaths {
    param(
        [string]$EffectiveOutDir,
        [string]$DbPath,
        [string]$NetDir,
        [string]$ExePath
    )
    $prepared = New-List
    $paths = New-List
    [void]$paths.Add('C:\ECG')
    [void]$paths.Add('C:\ECG\Tool')
    if (-not [string]::IsNullOrWhiteSpace($EffectiveOutDir)) { [void]$paths.Add($EffectiveOutDir) }
    if (Test-IsDrivePath $ExePath) { [void]$paths.Add((Split-Path -Parent $ExePath)) }
    foreach ($p in @($DbPath, $NetDir)) {
        if ([string]::IsNullOrWhiteSpace([string]$p)) { continue }
        if (Test-IsDrivePath $p) {
            [void]$paths.Add($p)
        }
        elseif (Test-IsUncPath $p) {
            $trim = Normalize-PathString $p
            if ($trim -match '^\\\\[^\\]+\\[^\\]+$') {
                Log ('UNC raiz detectado; não será criado: ' + $trim)
            }
            else {
                [void]$paths.Add($trim)
            }
        }
    }
    $seen = New-Map
    foreach ($p in @($paths)) {
        $n = Normalize-PathString ([string]$p)
        if ([string]::IsNullOrWhiteSpace($n)) { continue }
        if ($seen.ContainsKey($n)) { continue }
        $seen[$n] = $true
        $row = New-Map
        $row.Path = $n
        $row.Status = 'SKIPPED'
        $row.Message = 'Sem ação.'
        try {
            if (Test-Path -LiteralPath $n) {
                $row.Status = 'OK'
                $row.Message = 'Já existe.'
            }
            else {
                New-Item -ItemType Directory -Path $n -Force | Out-Null
                $row.Status = 'CREATED'
                $row.Message = 'Criado com sucesso.'
            }
        }
        catch {
            $row.Status = 'WARN'
            $row.Message = $_.Exception.Message
        }
        [void]$prepared.Add($row)
    }
    return $prepared
}

function Build-PrepareTextReport {
    param([array]$Rows)
    $lines = New-List
    [void]$lines.Add(($script:ToolName + ' ' + $script:ToolVersion))
    [void]$lines.Add(('Gerado em: ' + (Get-Date).ToString('o')))
    [void]$lines.Add('Modo: Prepare')
    [void]$lines.Add('')
    foreach ($row in @($Rows)) {
        [void]$lines.Add(('{0} | {1} | {2}' -f $row.Status, $row.Path, $row.Message))
    }
    return ($lines -join [Environment]::NewLine)
}

function Build-CompareReport {
    param(
        $Left,
        $Right,
        [string]$LeftPath,
        [string]$RightPath
    )
    $rows = New-List
    function Add-CompareRow {
        param([string]$Item,[object]$LeftValue,[object]$RightValue,[string]$Severity,[string]$Note)
        $row = New-Map
        $row.Item = $Item
        $row.Left = [string]$LeftValue
        $row.Right = [string]$RightValue
        $row.Match = ([string]$LeftValue -eq [string]$RightValue)
        $row.Severity = $Severity
        $row.Note = $Note
        [void]$rows.Add($row)
    }

    $leftCtx = Get-ReportValue -Object $Left -PathParts @('Context')
    $rightCtx = Get-ReportValue -Object $Right -PathParts @('Context')
    $leftAna = Get-ReportValue -Object $Left -PathParts @('Analysis')
    $rightAna = Get-ReportValue -Object $Right -PathParts @('Analysis')
    $leftEnv = Get-ReportValue -Object $Left -PathParts @('Environment')
    $rightEnv = Get-ReportValue -Object $Right -PathParts @('Environment')
    $leftIns = Get-ReportValue -Object $Left -PathParts @('Inspections')
    $rightIns = Get-ReportValue -Object $Right -PathParts @('Inspections')
    $leftTl = Get-ReportValue -Object $Left -PathParts @('TimelineSummary')
    $rightTl = Get-ReportValue -Object $Right -PathParts @('TimelineSummary')

    Add-CompareRow -Item 'Host' -LeftValue (Get-ReportValue $Left @('Meta','Host')) -RightValue (Get-ReportValue $Right @('Meta','Host')) -Severity 'INFO' -Note 'Identificação da estação.'
    Add-CompareRow -Item 'Role' -LeftValue (Get-ReportValue $Left @('Context','Role')) -RightValue (Get-ReportValue $Right @('Context','Role')) -Severity 'INFO' -Note 'Papel operacional.'
    Add-CompareRow -Item 'DesiredDbPath' -LeftValue (Get-ReportValue $Left @('Context','DesiredDbPath')) -RightValue (Get-ReportValue $Right @('Context','DesiredDbPath')) -Severity 'CRITICAL' -Note 'Caminho canônico esperado do banco.'
    Add-CompareRow -Item 'EffectiveDbPath' -LeftValue (Get-ReportValue $Left @('Context','EffectiveDbPath')) -RightValue (Get-ReportValue $Right @('Context','EffectiveDbPath')) -Severity 'CRITICAL' -Note 'Caminho efetivo observado.'
    Add-CompareRow -Item 'DesiredNetDir' -LeftValue (Get-ReportValue $Left @('Context','DesiredNetDir')) -RightValue (Get-ReportValue $Right @('Context','DesiredNetDir')) -Severity 'CRITICAL' -Note 'NETDIR esperado.'
    Add-CompareRow -Item 'EffectiveNetDir' -LeftValue (Get-ReportValue $Left @('Context','EffectiveNetDir')) -RightValue (Get-ReportValue $Right @('Context','EffectiveNetDir')) -Severity 'CRITICAL' -Note 'NETDIR efetivo.'
    Add-CompareRow -Item 'ExePath' -LeftValue (Get-ReportValue $Left @('Context','ExpectedExePath')) -RightValue (Get-ReportValue $Right @('Context','ExpectedExePath')) -Severity 'WARN' -Note 'Executável local do cliente.'

    $leftBde = Get-BdeNetDirMap -BdeNetDirState (Get-ReportValue $Left @('Environment','BdeNetDirState'))
    $rightBde = Get-BdeNetDirMap -BdeNetDirState (Get-ReportValue $Right @('Environment','BdeNetDirState'))
    foreach ($key in @('HKLM:\SOFTWARE\Borland\Database Engine\Settings\SYSTEM\INIT','HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT','HKCU:\Software\Borland\Database Engine\Settings\SYSTEM\INIT')) {
        Add-CompareRow -Item ('BDE ' + $key) -LeftValue (Get-MapValue $leftBde $key '') -RightValue (Get-MapValue $rightBde $key '') -Severity 'CRITICAL' -Note 'NETDIR por hive.'
    }

    Add-CompareRow -Item 'HeartWare Caminho Database' -LeftValue (Get-HeartWareDbPath -HeartWareSnapshots (Get-ReportValue $Left @('Environment','HeartWareRegistry'))) -RightValue (Get-HeartWareDbPath -HeartWareSnapshots (Get-ReportValue $Right @('Environment','HeartWareRegistry'))) -Severity 'WARN' -Note 'Configuração HeartWare por usuário.'
    Add-CompareRow -Item 'WriteProbe NetDir' -LeftValue (Get-ReportValue $Left @('Inspections','NetDir','WriteProbe','CanWrite')) -RightValue (Get-ReportValue $Right @('Inspections','NetDir','WriteProbe','CanWrite')) -Severity 'CRITICAL' -Note 'Permissão de escrita no NetDir.'
    Add-CompareRow -Item 'DbUnavailableSamples' -LeftValue (Get-ReportValue $Left @('TimelineSummary','DbUnavailableSamples')) -RightValue (Get-ReportValue $Right @('TimelineSummary','DbUnavailableSamples')) -Severity 'WARN' -Note 'Instabilidade temporal do banco.'
    Add-CompareRow -Item 'NetUnavailableSamples' -LeftValue (Get-ReportValue $Left @('TimelineSummary','NetDirUnavailableSamples')) -RightValue (Get-ReportValue $Right @('TimelineSummary','NetDirUnavailableSamples')) -Severity 'WARN' -Note 'Instabilidade temporal do NetDir.'
    Add-CompareRow -Item 'PeakLockCount' -LeftValue (Get-ReportValue $Left @('TimelineSummary','PeakLockCount')) -RightValue (Get-ReportValue $Right @('TimelineSummary','PeakLockCount')) -Severity 'INFO' -Note 'Pico de locks observado.'
    Add-CompareRow -Item 'PrimaryCategory' -LeftValue (Get-ReportValue $Left @('Analysis','PrimaryCategory')) -RightValue (Get-ReportValue $Right @('Analysis','PrimaryCategory')) -Severity 'INFO' -Note 'Categoria principal do laudo.'

    $criticalMismatch = 0
    $warnMismatch = 0
    foreach ($row in @($rows)) {
        if (-not $row.Match) {
            if ($row.Severity -eq 'CRITICAL') { $criticalMismatch++ }
            elseif ($row.Severity -eq 'WARN') { $warnMismatch++ }
        }
    }

    $status = 'CONVERGENTE'
    $apt = 'APTO_PARA_PILOTO'
    if ($criticalMismatch -gt 0) {
        $status = 'DRIFT_CRITICO'
        $apt = 'NAO_APTO'
    }
    elseif ($warnMismatch -gt 0) {
        $status = 'DRIFT_LEVE'
    }

    $findings = New-List
    $recommendations = New-List
    foreach ($row in @($rows)) {
        if (-not $row.Match) {
            [void]$findings.Add(($row.Item + ' diverge entre as estações.'))
            if ($row.Severity -eq 'CRITICAL') {
                [void]$recommendations.Add(('Padronizar imediatamente: ' + $row.Item))
            }
        }
    }
    if ($findings.Count -eq 0) {
        [void]$findings.Add('Os principais vetores de configuração convergiram entre os dois laudos.')
    }
    if ($recommendations.Count -eq 0) {
        [void]$recommendations.Add('Manter o perfil atual e seguir para validação funcional do ECGv6.')
    }

    $compare = New-Map
    $compare.Meta = New-Map
    $compare.Meta.ToolName = $script:ToolName
    $compare.Meta.ToolVersion = $script:ToolVersion
    $compare.Meta.GeneratedAt = (Get-Date).ToString('o')
    $compare.Meta.Mode = 'Compare'
    $compare.Meta.LeftReport = $LeftPath
    $compare.Meta.RightReport = $RightPath
    $compare.Status = $status
    $compare.Readiness = $apt
    $compare.CriticalMismatchCount = $criticalMismatch
    $compare.WarnMismatchCount = $warnMismatch
    $compare.Rows = $rows
    $compare.Findings = $findings
    $compare.Recommendations = $recommendations
    return $compare
}

function Build-CompareTextReport {
    param($Compare)
    $lines = New-List
    [void]$lines.Add(($script:ToolName + ' ' + $script:ToolVersion))
    [void]$lines.Add(('Gerado em: ' + $Compare.Meta.GeneratedAt))
    [void]$lines.Add('Modo: Compare')
    [void]$lines.Add(('Status: ' + $Compare.Status))
    [void]$lines.Add(('Readiness: ' + $Compare.Readiness))
    [void]$lines.Add(('LeftReport : ' + $Compare.Meta.LeftReport))
    [void]$lines.Add(('RightReport: ' + $Compare.Meta.RightReport))
    [void]$lines.Add('')
    [void]$lines.Add('DIVERGENCIAS')
    foreach ($row in @($Compare.Rows)) {
        $marker = 'OK'
        if (-not $row.Match) { $marker = 'DIFF' }
        [void]$lines.Add(('{0} | {1} | L={2} | R={3} | {4}' -f $marker, $row.Item, $row.Left, $row.Right, $row.Note))
    }
    [void]$lines.Add('')
    [void]$lines.Add('ACHADOS')
    foreach ($f in @($Compare.Findings)) { [void]$lines.Add(('  - ' + $f)) }
    [void]$lines.Add('')
    [void]$lines.Add('RECOMENDACOES')
    foreach ($r in @($Compare.Recommendations)) { [void]$lines.Add(('  - ' + $r)) }
    return ($lines -join [Environment]::NewLine)
}

function Build-CompareHtmlReport {
    param($Compare)
    $rows = ''
    foreach ($row in @($Compare.Rows)) {
        $klass = 'ok'
        if (-not $row.Match -and $row.Severity -eq 'CRITICAL') { $klass = 'fail' }
        elseif (-not $row.Match) { $klass = 'warn' }
        $statusText = 'OK'
        if (-not $row.Match) { $statusText = 'DIFF' }
        $rows += ('<tr><td>' + (Escape-Html ([string]$row.Item)) + '</td><td>' + (Escape-Html ([string]$row.Left)) + '</td><td>' + (Escape-Html ([string]$row.Right)) + '</td><td>' + (Escape-Html ([string]$row.Note)) + '</td><td><span class="badge ' + $klass + '">' + (Escape-Html ([string]$statusText)) + '</span></td></tr>')
    }
    $findingsHtml = ''
    foreach ($f in @($Compare.Findings)) { $findingsHtml += ('<li>' + (Escape-Html ([string]$f)) + '</li>') }
    $recoHtml = ''
    foreach ($r in @($Compare.Recommendations)) { $recoHtml += ('<li>' + (Escape-Html ([string]$r)) + '</li>') }
    $statusClass = 'ok'
    if ([string]$Compare.Status -eq 'DRIFT_CRITICO') { $statusClass = 'fail' }
    elseif ([string]$Compare.Status -eq 'DRIFT_LEVE') { $statusClass = 'warn' }
    return @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECGv6 FieldKit Compare</title>
<style>
body { font-family: Segoe UI, Tahoma, Arial; margin: 24px; color: #111827; }
.card { border: 1px solid #d1d5db; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
.badge { display: inline-block; padding: 6px 10px; border-radius: 999px; font-weight: 700; color: #fff; }
.ok { background: #059669; }
.warn { background: #d97706; }
.fail { background: #dc2626; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #e5e7eb; padding: 8px; text-align: left; vertical-align: top; }
th { background: #f9fafb; }
</style>
</head>
<body>
<h1>ECGv6 FieldKit Compare</h1>
<div class="card">
  <div class="badge $statusClass">$([string](Escape-Html ([string]$Compare.Status)))</div>
  <p><strong>Readiness:</strong> $([string](Escape-Html ([string]$Compare.Readiness)))</p>
  <p><strong>Left:</strong> $([string](Escape-Html ([string]$Compare.Meta.LeftReport)))</p>
  <p><strong>Right:</strong> $([string](Escape-Html ([string]$Compare.Meta.RightReport)))</p>
</div>
<div class="card"><h2>Achados</h2><ul>$findingsHtml</ul><h2>Recomendações</h2><ul>$recoHtml</ul></div>
<div class="card"><h2>Comparação</h2><table><tr><th>Item</th><th>Left</th><th>Right</th><th>Nota</th><th>Status</th></tr>$rows</table></div>
</body>
</html>
"@
}

function Invoke-RollbackReg {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { throw 'Informe -RollbackFile.' }
    if (-not (Test-Path -LiteralPath $Path)) { throw ('Arquivo de rollback não encontrado: ' + $Path) }
    if (-not (Test-IsAdmin)) {
        if (-not $Elevated) {
            Log 'Rollback solicitado sem elevação. Solicitando UAC...' 'STEP'
            $argString = Get-RelaunchArgumentString -ForElevation
            Start-Process -FilePath 'powershell.exe' -ArgumentList $argString -Verb RunAs -Wait
            return $false
        }
        throw 'Rollback requer privilégios administrativos.'
    }
    $proc = Start-Process -FilePath 'reg.exe' -ArgumentList @('import', $Path) -Wait -PassThru -WindowStyle Hidden
    if ($proc.ExitCode -ne 0) { throw ('reg import retornou código ' + [string]$proc.ExitCode) }
    return $true
}

function Resolve-DesiredState {
    param(
        $Profile,
        $EnvState,
        [array]$HeartWareSnapshots,
        [array]$Shares,
        [array]$NetworkConnections,
        [string]$ExpectedDbPathParam,
        [string]$ExpectedNetDirParam,
        [string]$ExpectedExePathParam
    )

    $dbCandidates = New-List
    $directCandidates = @(
        $ExpectedDbPathParam,
        $Profile['ExpectedDbPath'],
        $EnvState.Process,
        $EnvState.User,
        $EnvState.Machine
    )
    foreach ($candidate in $directCandidates) {
        if (-not [string]::IsNullOrWhiteSpace([string]$candidate)) {
            [void]$dbCandidates.Add((Normalize-PathString ([string]$candidate)))
        }
    }

    foreach ($snap in @($HeartWareSnapshots)) {
        if (-not $snap.Exists) { continue }
        foreach ($entry in $snap.Values.GetEnumerator()) {
            if ($entry.Value -is [string] -and $entry.Value -match '^[A-Za-z]:\\|^\\\\') {
                [void]$dbCandidates.Add((Normalize-PathString ([string]$entry.Value)))
            }
        }
    }

    $preferredDb = $null
    foreach ($candidate in @($dbCandidates)) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
        $converted = Convert-DrivePathToUnc -Path $candidate -NetworkConnections $NetworkConnections
        if (Test-IsUncPath $converted) {
            $preferredDb = $converted
            break
        }
    }

    if ($null -eq $preferredDb) {
        foreach ($candidate in @($dbCandidates)) {
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $shareResolution = Resolve-ShareForLocalPath -LocalPath $candidate -Shares $Shares -ComputerName $script:HostName
            if ($shareResolution -and -not [string]::IsNullOrWhiteSpace($shareResolution.UncPath)) {
                $preferredDb = $shareResolution.UncPath
                break
            }
        }
    }

    if ($null -eq $preferredDb) {
        foreach ($candidate in @($dbCandidates)) {
            if ([string]::IsNullOrWhiteSpace($candidate)) { continue }
            $preferredDb = $candidate
            break
        }
    }

    $preferredNetDir = Coalesce-NonEmpty @($ExpectedNetDirParam, $Profile['ExpectedNetDir'])
    if ([string]::IsNullOrWhiteSpace([string]$preferredNetDir) -and -not [string]::IsNullOrWhiteSpace([string]$preferredDb)) {
        $preferredNetDir = Join-Path (Normalize-PathString ([string]$preferredDb)) 'NetDir'
    }

    $preferredExe = Coalesce-NonEmpty @($ExpectedExePathParam, $Profile['ExpectedExePath'], 'C:\HW\ECG\ECGV6.exe')

    $obj = New-Map
    $obj.PreferredDbPath = Normalize-PathString ([string]$preferredDb)
    $obj.PreferredNetDir = Normalize-PathString ([string]$preferredNetDir)
    $obj.PreferredExePath = Normalize-PathString ([string]$preferredExe)
    $obj.SetMachineHwPath = (($SetMachineHwPath -eq $true) -or (([string]$Profile['SetMachineHwPath']).ToUpperInvariant() -eq 'TRUE'))
    return $obj
}

function Get-CurrentState {
    param(
        $EnvState,
        [array]$BdeNetDirState,
        [array]$HeartWareSnapshots,
        [array]$Shares,
        [array]$NetworkConnections
    )

    $effectiveDb = Coalesce-NonEmpty @($EnvState.Process, $EnvState.User, $EnvState.Machine)
    if ([string]::IsNullOrWhiteSpace([string]$effectiveDb)) {
        foreach ($snap in @($HeartWareSnapshots)) {
            if (-not $snap.Exists) { continue }
            foreach ($entry in $snap.Values.GetEnumerator()) {
                if ($entry.Value -is [string] -and $entry.Value -match '^[A-Za-z]:\\|^\\\\') {
                    $effectiveDb = $entry.Value
                    break
                }
            }
            if (-not [string]::IsNullOrWhiteSpace([string]$effectiveDb)) { break }
        }
    }

    $effectiveDb = Normalize-PathString ([string](Convert-DrivePathToUnc -Path $effectiveDb -NetworkConnections $NetworkConnections))
    if (-not (Test-IsUncPath $effectiveDb) -and -not [string]::IsNullOrWhiteSpace([string]$effectiveDb)) {
        $shareResolution = Resolve-ShareForLocalPath -LocalPath $effectiveDb -Shares $Shares -ComputerName $script:HostName
        if ($shareResolution -and -not [string]::IsNullOrWhiteSpace($shareResolution.UncPath)) {
            $effectiveDb = Normalize-PathString $shareResolution.UncPath
        }
    }

    $effectiveNetDir = $null
    foreach ($row in @($BdeNetDirState)) {
        if (-not [string]::IsNullOrWhiteSpace([string]$row.CurrentValue)) {
            $effectiveNetDir = Normalize-PathString ([string]$row.CurrentValue)
            break
        }
    }
    if ([string]::IsNullOrWhiteSpace([string]$effectiveNetDir) -and -not [string]::IsNullOrWhiteSpace([string]$effectiveDb)) {
        $effectiveNetDir = Normalize-PathString (Join-Path $effectiveDb 'NetDir')
    }

    $obj = New-Map
    $obj.EffectiveDbPath = $effectiveDb
    $obj.EffectiveNetDir = $effectiveNetDir
    return $obj
}

function New-ChangePlanItem {
    param(
        [string]$Key,
        [string]$Action,
        [string]$Target,
        [string]$CurrentValue,
        [string]$DesiredValue,
        [string]$Reason,
        [string]$Safety = 'SAFE'
    )
    $row = New-Map
    $row.Key = $Key
    $row.Action = $Action
    $row.Target = $Target
    $row.CurrentValue = $CurrentValue
    $row.DesiredValue = $DesiredValue
    $row.Reason = $Reason
    $row.Safety = $Safety
    return $row
}

function Build-ChangePlan {
    param(
        $DesiredState,
        $CurrentState,
        $EnvState,
        [array]$BdeNetDirState
    )

    $plan = New-List

    if (-not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredNetDir)) {
        foreach ($root in @($BdeNetDirState)) {
            $current = Normalize-PathString ([string]$root.CurrentValue)
            $shouldManageRoot = $true
            # ===== MODIFICAÇÃO: criar HKCU se não existir =====
            if (($root.RegistryPath -like 'HKCU:*') -and (-not $root.Exists) -and [string]::IsNullOrWhiteSpace([string]$root.CurrentValue)) {
                $shouldManageRoot = $true   # antes era $false
            }
            if ($shouldManageRoot -and ($current -ne $DesiredState.PreferredNetDir)) {
                $reason = 'Padronizar NETDIR do BDE para caminho único e determinístico.'
                [void]$plan.Add((New-ChangePlanItem -Key ('BDE_NETDIR|' + $root.RegistryPath) -Action 'SET_BDE_NETDIR' -Target $root.RegistryPath -CurrentValue $current -DesiredValue $DesiredState.PreferredNetDir -Reason $reason -Safety 'SAFE'))
            }
        }
    }

    $shouldSetMachineHw = $false
    if ($DesiredState.SetMachineHwPath) {
        $shouldSetMachineHw = $true
    }
    elseif (-not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredDbPath) -and [string]::IsNullOrWhiteSpace([string]$EnvState.Process) -and [string]::IsNullOrWhiteSpace([string]$EnvState.User) -and [string]::IsNullOrWhiteSpace([string]$EnvState.Machine)) {
        $shouldSetMachineHw = $true
    }

    if ($shouldSetMachineHw -and -not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredDbPath)) {
        $currentMachine = Normalize-PathString ([string]$EnvState.Machine)
        if ($currentMachine -ne $DesiredState.PreferredDbPath) {
            [void]$plan.Add((New-ChangePlanItem -Key 'ENV_MACHINE_HW_CAMINHO_DB' -Action 'SET_MACHINE_ENV' -Target 'HKLM:Environment/HW_CAMINHO_DB' -CurrentValue $currentMachine -DesiredValue $DesiredState.PreferredDbPath -Reason 'Padronizar HW_CAMINHO_DB em nível de máquina para todos os usuários.' -Safety 'SAFE'))
        }
    }

    if ($CreateMissingDirs -and -not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredNetDir)) {
        if (-not (Test-Path -LiteralPath $DesiredState.PreferredNetDir)) {
            $safety = 'CAUTION'
            if (-not (Test-IsUncPath $DesiredState.PreferredNetDir)) {
                $safety = 'SAFE'
            }
            [void]$plan.Add((New-ChangePlanItem -Key 'CREATE_NETDIR' -Action 'CREATE_DIRECTORY' -Target $DesiredState.PreferredNetDir -CurrentValue '' -DesiredValue $DesiredState.PreferredNetDir -Reason 'Criar NETDIR ausente.' -Safety $safety))
        }
    }

    return $plan
}

function Quote-Arg {
    param([string]$Text)
    if ($null -eq $Text) { return '""' }
    $escaped = [string]$Text -replace '"', '`"'
    if ($escaped -match '\s|"') {
        return ('"' + $escaped + '"')
    }
    return $escaped
}

function Get-RelaunchArgumentString {
    param([switch]$ForElevation)
    $parts = New-List
    [void]$parts.Add('-NoProfile')
    [void]$parts.Add('-ExecutionPolicy')
    [void]$parts.Add('Bypass')
    [void]$parts.Add('-File')
    [void]$parts.Add((Quote-Arg $script:ScriptPath))

    foreach ($entry in $PSBoundParameters.GetEnumerator()) {
        $name = [string]$entry.Key
        if ($name -eq 'Elevated') { continue }
        $value = $entry.Value
        if ($value -is [switch]) {
            if ($value.IsPresent) {
                [void]$parts.Add(('-' + $name))
            }
        }
        else {
            [void]$parts.Add(('-' + $name))
            [void]$parts.Add((Quote-Arg ([string]$value)))
        }
    }

    if ($ForElevation) {
        [void]$parts.Add('-Elevated')
    }

    return ($parts -join ' ')
}

function Ensure-AdminForFixes {
    param([array]$Plan)
    if ($Mode -ne 'Fix') { return $true }
    if (($null -eq $Plan) -or ($Plan.Count -eq 0)) { return $true }
    if (Test-IsAdmin) { return $true }
    if ($Elevated) { return $false }

    Log 'Correções pendentes detectadas e o processo atual não está elevado. Solicitando elevação...' 'STEP'
    $argString = Get-RelaunchArgumentString -ForElevation
    try {
        Start-Process -FilePath 'powershell.exe' -ArgumentList $argString -Verb RunAs -Wait
        return $false
    }
    catch {
        Fail-Note ('Falha ao solicitar elevação: ' + $_.Exception.Message)
        return $false
    }
}

function Ensure-RegistryPath {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -Path $Path -Force | Out-Null
    }
}

function Backup-RegistryBranch {
    param(
        [string]$PsPath,
        [string]$BackupDir,
        [string]$NameHint
    )
    try {
        $native = $PsPath -replace '^HKLM:\', 'HKLM\' -replace '^HKCU:\', 'HKCU\'
        $safeName = ($NameHint -replace '[^A-Za-z0-9._-]', '_')
        $file = Join-Path $BackupDir ($safeName + '.reg')
        $query = Start-Process -FilePath 'reg.exe' -ArgumentList @('query', $native) -Wait -PassThru -WindowStyle Hidden
        if ($query.ExitCode -ne 0) { return '__REGKEY_NOT_FOUND__' }
        $export = Start-Process -FilePath 'reg.exe' -ArgumentList @('export', $native, $file, '/y') -Wait -PassThru -WindowStyle Hidden
        if ($export.ExitCode -ne 0) { return $null }
        if (Test-Path -LiteralPath $file) { return $file }
        return $null
    }
    catch {
        return $null
    }
}

function Apply-ChangePlan {
    param(
        [array]$Plan,
        [string]$BackupDir
    )

    $results = New-List
    foreach ($item in @($Plan)) {
        $row = New-Map
        $row.Key = $item.Key
        $row.Action = $item.Action
        $row.Target = $item.Target
        $row.Status = 'SKIPPED'
        $row.Message = 'Sem ação.'
        $row.Backup = $null

        try {
            switch ($item.Action) {
                'SET_BDE_NETDIR' {
                    $targetExisted = [bool](Test-Path -LiteralPath $item.Target)
                    $backupFile = Backup-RegistryBranch -PsPath $item.Target -BackupDir $BackupDir -NameHint ('backup_' + ($item.Target -replace '[:\ ]','_'))
                    if ($backupFile -eq '__REGKEY_NOT_FOUND__') {
                        $row.Backup = 'N/A - chave inexistente antes da correção'
                    }
                    else {
                        $row.Backup = $backupFile
                    }
                    Ensure-RegistryPath -Path $item.Target
                    if ($PSCmdlet.ShouldProcess($item.Target, 'Definir NETDIR do BDE')) {
                        Set-ItemProperty -Path $item.Target -Name 'NETDIR' -Value $item.DesiredValue -Type String -Force
                    }
                    $row.Status = 'APPLIED'
                    if ($backupFile -and $backupFile -ne '__REGKEY_NOT_FOUND__') {
                        $row.Message = 'NETDIR gravado com sucesso. Backup .reg exportado.'
                    }
                    elseif (-not $targetExisted) {
                        $row.Message = 'NETDIR gravado com sucesso. Backup .reg não aplicável porque a chave não existia.'
                    }
                    else {
                        $row.Message = 'NETDIR gravado com sucesso. Backup .reg não pôde ser evidenciado.'
                    }
                    Add-ListItem -List $script:AppliedChanges -Value $row
                }
                'SET_MACHINE_ENV' {
                    if ($PSCmdlet.ShouldProcess($item.Target, 'Definir variável de ambiente HW_CAMINHO_DB')) {
                        [Environment]::SetEnvironmentVariable('HW_CAMINHO_DB', $item.DesiredValue, 'Machine')
                    }
                    $row.Status = 'APPLIED'
                    $row.Message = 'HW_CAMINHO_DB em nível de máquina atualizado.'
                    Add-ListItem -List $script:AppliedChanges -Value $row
                }
                'CREATE_DIRECTORY' {
                    if ($PSCmdlet.ShouldProcess($item.Target, 'Criar diretório')) {
                        New-Item -Path $item.Target -ItemType Directory -Force | Out-Null
                    }
                    $row.Status = 'APPLIED'
                    $row.Message = 'Diretório criado ou já existente.'
                    Add-ListItem -List $script:AppliedChanges -Value $row
                }
                default {
                    $row.Status = 'SKIPPED'
                    $row.Message = 'Ação não reconhecida.'
                }
            }
        }
        catch {
            $row.Status = 'ERROR'
            $row.Message = $_.Exception.Message
            Add-ListItem -List $script:ErrorsFound -Value ($item.Action + ': ' + $_.Exception.Message)
        }

        [void]$results.Add($row)
    }

    return $results
}

# ----- NOVA FUNÇÃO: CORRIGIR IDAPI32.CFG -----
function Repair-IdapiCfgNetDir {
    param(
        [string]$DesiredNetDir,
        [switch]$Force
    )
    $result = New-Map
    $result.Modified = $false
    $result.Verified = $false
    $result.FilePath = $null
    $result.BackupPath = $null
    $result.OldValue = $null
    $result.NewValue = $DesiredNetDir
    $result.Status = 'SKIPPED'
    $result.Message = 'Sem ação.'

    if ([string]::IsNullOrWhiteSpace($DesiredNetDir)) {
        $result.Message = 'DesiredNetDir vazio.'
        return $result
    }

    $pathsToTry = @(
        'C:\Program Files (x86)\Common Files\Borland Shared\BDE\IDAPI32.CFG',
        'C:\Program Files\Common Files\Borland Shared\BDE\IDAPI32.CFG',
        'C:\Arquivos de Programas\Arquivos Comuns\Borland Shared\BDE\IDAPI32.CFG'
    )

    foreach ($path in $pathsToTry) {
        if (-not (Test-Path -LiteralPath $path)) { continue }
        try {
            $bytes = [System.IO.File]::ReadAllBytes($path)
            $enc = [System.Text.Encoding]::Default
            $content = $enc.GetString($bytes)
            $result.FilePath = $path

            $candidateValues = @(
                'C:\HW\Database\NetDir',
                'C:\HW\DATABASE\NETDIR',
                'C:\HW\Database\NETDIR',
                'C:\HW\database\NetDir'
            )
            $oldNetDir = $null
            foreach ($candidate in $candidateValues) {
                if ($content.IndexOf($candidate, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) {
                    $oldNetDir = $candidate
                    break
                }
            }

            if ([string]::IsNullOrWhiteSpace([string]$oldNetDir)) {
                $m = [regex]::Match($content, '(?is)NET\s*DIR[^A-Za-z0-9/:_\-]{1,32}([A-Za-z]:\\[^\x00\r\n]+|\\\\[^\x00\r\n]+)')
                if ($m.Success) {
                    $oldNetDir = [string]$m.Groups[1].Value
                }
            }

            $result.OldValue = $oldNetDir
            if ([string]::IsNullOrWhiteSpace([string]$oldNetDir)) {
                $result.Status = 'SKIPPED'
                $result.Message = 'IDAPI32.CFG sem NET DIR identificável.'
                return $result
            }

            if ((Normalize-PathString $oldNetDir) -eq (Normalize-PathString $DesiredNetDir)) {
                $result.Status = 'OK'
                $result.Verified = $true
                $result.Message = 'IDAPI32.CFG já estava alinhado com o UNC desejado.'
                return $result
            }

            $backup = $path + '.bak'
            Copy-Item -LiteralPath $path -Destination $backup -Force
            $newContent = $content.Replace($oldNetDir, $DesiredNetDir)
            [System.IO.File]::WriteAllBytes($path, $enc.GetBytes($newContent))

            $verifyBytes = [System.IO.File]::ReadAllBytes($path)
            $verifyContent = $enc.GetString($verifyBytes)
            $hasDesired = ($verifyContent.IndexOf($DesiredNetDir, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)
            $hasOld = ($verifyContent.IndexOf($oldNetDir, [System.StringComparison]::OrdinalIgnoreCase) -ge 0)

            $result.Modified = $true
            $result.BackupPath = $backup
            $result.Verified = ($hasDesired -and (-not $hasOld))
            if ($result.Verified) {
                $result.Status = 'APPLIED'
                $result.Message = 'IDAPI32.CFG corrigido e validado com backup .bak.'
                Log ('IDAPI32.CFG corrigido e validado: ' + $path + ' | antigo=' + $oldNetDir + ' | novo=' + $DesiredNetDir) 'STEP'
            }
            else {
                $result.Status = 'WARN'
                $result.Message = 'IDAPI32.CFG alterado, mas a validação pós-gravação não confirmou a substituição completa.'
                Warn ('IDAPI32.CFG alterado sem validação completa: ' + $path)
            }
            return $result
        }
        catch {
            $result.Status = 'WARN'
            $result.Message = $_.Exception.Message
            Warn ('Não foi possível corrigir IDAPI32.CFG em ' + $path + ' : ' + $_.Exception.Message)
            return $result
        }
    }

    $result.Message = 'IDAPI32.CFG não encontrado nas trilhas esperadas.'
    return $result
}

function Get-Sample {
    param(
        [string]$DbPath,
        [string]$NetDirPath
    )

    $sample = New-Map
    $sample.Timestamp = (Get-Date).ToString('o')
    $sample.DbAccessible = $false
    $sample.NetDirAccessible = $false
    $sample.LockCount = 0
    $sample.Error = $null

    try {
        if (-not [string]::IsNullOrWhiteSpace([string]$DbPath)) {
            $sample.DbAccessible = [bool](Test-Path -LiteralPath $DbPath)
        }
        if (-not [string]::IsNullOrWhiteSpace([string]$NetDirPath)) {
            $sample.NetDirAccessible = [bool](Test-Path -LiteralPath $NetDirPath)
            if ($sample.NetDirAccessible) {
                $locks = Get-LockInventory -NetDirPath $NetDirPath
                $sample.LockCount = [int]$locks.Count
            }
        }
    }
    catch {
        $sample.Error = $_.Exception.Message
    }

    return $sample
}

function Collect-Timeline {
    param(
        [string]$DbPath,
        [string]$NetDirPath,
        [int]$Minutes,
        [int]$IntervalSeconds
    )

    $timeline = New-List
    if ($Minutes -le 0) { return $timeline }
    $targetSamples = [Math]::Max(1, [Math]::Floor(($Minutes * 60) / $IntervalSeconds))
    for ($i = 1; $i -le $targetSamples; $i++) {
        $sample = Get-Sample -DbPath $DbPath -NetDirPath $NetDirPath
        [void]$timeline.Add($sample)
        Log ('Amostra ' + $i + '/' + $targetSamples + ' | DB=' + [string]$sample.DbAccessible + ' | NetDir=' + [string]$sample.NetDirAccessible + ' | Locks=' + [string]$sample.LockCount)
        if ($i -lt $targetSamples) {
            Start-Sleep -Seconds $IntervalSeconds
        }
    }
    return $timeline
}

function Get-TimelineSummary {
    param([array]$Timeline)
    $summary = New-Map
    $summary.SampleCount = @($Timeline).Count
    $summary.DbUnavailableSamples = 0
    $summary.NetDirUnavailableSamples = 0
    $summary.PeakLockCount = 0
    foreach ($sample in @($Timeline)) {
        if ($sample.DbAccessible -ne $true) { $summary.DbUnavailableSamples++ }
        if ($sample.NetDirAccessible -ne $true) { $summary.NetDirUnavailableSamples++ }
        if ([int]$sample.LockCount -gt [int]$summary.PeakLockCount) {
            $summary.PeakLockCount = [int]$sample.LockCount
        }
    }
    return $summary
}

function Get-ShareHealthState {
    param(
        [array]$Shares,
        [string]$DesiredDbPath
    )
    $obj = New-Map
    $obj.HasLocalHwShare = $false
    $obj.HwSharePath = $null
    $obj.HwShareMatchesDb = $null

    foreach ($share in @($Shares)) {
        if ([string]$share.Name -ieq 'HW') {
            $obj.HasLocalHwShare = $true
            $obj.HwSharePath = Normalize-PathString ([string]$share.Path)
            break
        }
    }

    if ($obj.HasLocalHwShare -and -not [string]::IsNullOrWhiteSpace([string]$DesiredDbPath) -and -not (Test-IsUncPath $DesiredDbPath)) {
        $obj.HwShareMatchesDb = ([string]$obj.HwSharePath).ToUpperInvariant() -eq ([string](Normalize-PathString $DesiredDbPath)).ToUpperInvariant()
    }

    return $obj
}

function Get-ConsistencyState {
    param(
        $DesiredState,
        $CurrentState,
        [array]$BdeNetDirState,
        $EnvState,
        $DbInspection,
        $NetInspection,
        $WriteProbeResult,
        $TimelineSummary,
        $ShareHealth,
        [array]$IdapiCfgState,
        [array]$HeartWareSnapshots
    )

    $scores = New-Map
    $scores.CONFIG = 0
    $scores.SHARE = 0
    $scores.PERMISSION = 0
    $scores.LOCK = 0
    $scores.SOFTWARE = 0

    $findings = New-List
    $recommendations = New-List

    if ([string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredDbPath)) {
        $scores.CONFIG += 3
        [void]$findings.Add('Caminho desejado do banco não pôde ser determinado de forma inequívoca.')
        [void]$recommendations.Add('Informar explicitamente o caminho UNC canônico do banco no parâmetro -ExpectedDbPath ou no profile.')
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$CurrentState.EffectiveDbPath) -and -not (Test-IsUncPath $CurrentState.EffectiveDbPath)) {
        $scores.CONFIG += 4
        [void]$findings.Add('O caminho efetivo do banco não está em UNC direto.')
        [void]$recommendations.Add('Eliminar unidade mapeada e padronizar HW_CAMINHO_DB em UNC direto.')
    }

    if (-not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredNetDir)) {
        foreach ($row in @($BdeNetDirState)) {
            $current = Normalize-PathString ([string]$row.CurrentValue)
            if ($current -ne $DesiredState.PreferredNetDir) {
                $scores.CONFIG += 2
                [void]$findings.Add(('NETDIR divergente em ' + $row.RegistryPath + '.'))
            }
        }
    }


    $heartWareDb = Get-HeartWareDbPath -HeartWareSnapshots $HeartWareSnapshots
    if (-not [string]::IsNullOrWhiteSpace([string]$heartWareDb) -and $DesiredState.PreferredDbPath -and ((Normalize-PathString $heartWareDb) -ne $DesiredState.PreferredDbPath)) {
        $scores.CONFIG += 2
        [void]$findings.Add('HeartWare aponta para caminho de banco divergente do desejado.')
        [void]$recommendations.Add('Padronizar HeartWare\ECGV6\Geral\Caminho Database com o UNC canônico.')
    }

    foreach ($cfg in @($IdapiCfgState)) {
        if (-not $cfg.Exists) { continue }
        $extract = [string]$cfg.Extract
        if ([string]::IsNullOrWhiteSpace($extract)) { continue }
        $normalizedExtract = ($extract -replace [char]0, ' ')
        if (($normalizedExtract -match 'NET\s*DIR|NETDIR') -and (-not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredNetDir))) {
            if ($normalizedExtract -notmatch [regex]::Escape($DesiredState.PreferredNetDir)) {
                $scores.CONFIG += 2
                [void]$findings.Add(('IDAPI32.CFG contém rastro de NETDIR divergente em ' + [string]$cfg.Path + '.'))
                [void]$recommendations.Add('Revisar BDE Administrator / IDAPI32.CFG para remover referência antiga/local de NETDIR.')
            }
        }
    }

    if ($DbInspection.Exists -ne $true) {
        $scores.SHARE += 5
        [void]$findings.Add('Banco inacessível no caminho efetivo da rodada.')
        [void]$recommendations.Add('Validar share SMB, nome UNC, serviço Server e reachability da VM XP.')
    }

    if ($NetInspection.Exists -ne $true) {
        $scores.SHARE += 5
        [void]$findings.Add('NetDir inacessível no caminho efetivo da rodada.')
        [void]$recommendations.Add('Garantir NetDir único e acessível via UNC em todos os clientes.')
    }

    if ($WriteProbeResult -and $WriteProbeResult.Attempted -and $WriteProbeResult.CanWrite -ne $true) {
        $scores.PERMISSION += 5
        [void]$findings.Add('A estação não conseguiu gravar no NetDir.')
        [void]$recommendations.Add('Revisar permissões de compartilhamento e NTFS do NetDir.')
    }

    if ($TimelineSummary.DbUnavailableSamples -gt 0 -or $TimelineSummary.NetDirUnavailableSamples -gt 0) {
        $scores.SHARE += [Math]::Min(4, ($TimelineSummary.DbUnavailableSamples + $TimelineSummary.NetDirUnavailableSamples))
        [void]$findings.Add('Houve instabilidade temporal de acesso durante a janela observada.')
    }

    if ($TimelineSummary.PeakLockCount -gt 0 -and $NetInspection.Exists -eq $true) {
        $scores.LOCK += [Math]::Min(4, [int]$TimelineSummary.PeakLockCount)
        [void]$findings.Add(('Arquivos de lock/controle observados no NetDir. Pico=' + [string]$TimelineSummary.PeakLockCount))
        [void]$recommendations.Add('Garantir NetDir único e evitar divergência entre estação executante e viewers.')
    }

    if ($ShareHealth.HasLocalHwShare -eq $true -and -not [string]::IsNullOrWhiteSpace([string]$ShareHealth.HwSharePath) -and -not [string]::IsNullOrWhiteSpace([string]$DesiredState.PreferredDbPath) -and -not (Test-IsUncPath $DesiredState.PreferredDbPath)) {
        if ($ShareHealth.HwShareMatchesDb -eq $false) {
            $scores.CONFIG += 3
            [void]$findings.Add('O share local HW não aponta para o mesmo caminho do banco esperado.')
        }
    }

    $primaryKey = 'SOFTWARE'
    $maxScore = -1
    foreach ($entry in $scores.GetEnumerator()) {
        if ([int]$entry.Value -gt $maxScore) {
            $maxScore = [int]$entry.Value
            $primaryKey = [string]$entry.Key
        }
    }

    if ($maxScore -le 0) {
        $primaryKey = 'OK'
    }

    $status = 'OK'
    if ($primaryKey -ne 'OK') {
        if ($maxScore -ge 6) { $status = 'FALHA' }
        else { $status = 'ALERTA' }
    }

    $confidence = 'Média'
    if ($maxScore -ge 8) { $confidence = 'Alta' }
    elseif ($maxScore -le 3) { $confidence = 'Baixa' }

    $causeLabel = 'Ambiente íntegro nesta rodada.'
    switch ($primaryKey) {
        'CONFIG' { $causeLabel = 'Configuração/BDE/UNC inconsistente' }
        'SHARE' { $causeLabel = 'Compartilhamento/rede/SMB instável ou indisponível' }
        'PERMISSION' { $causeLabel = 'Permissão insuficiente no NetDir' }
        'LOCK' { $causeLabel = 'Contenção/lock/concorrência do BDE' }
        'SOFTWARE' { $causeLabel = 'Software/arquivo sem evidência conclusiva de infraestrutura' }
        'OK' { $causeLabel = 'Infraestrutura e configuração consistentes na janela observada' }
    }

    $obj = New-Map
    $obj.Status = $status
    $obj.PrimaryCategory = $primaryKey
    $obj.PrimaryCause = $causeLabel
    $obj.Confidence = $confidence
    $obj.Scores = $scores
    $obj.Findings = $findings
    $obj.Recommendations = $recommendations
    return $obj
}

function Escape-Html {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $value = [string]$Text
    $value = $value.Replace('&', '&amp;')
    $value = $value.Replace('<', '&lt;')
    $value = $value.Replace('>', '&gt;')
    $value = $value.Replace('"', '&quot;')
    return $value
}

function Escape-JsonString {
    param([string]$Text)
    if ($null -eq $Text) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Text.ToCharArray()) {
        switch ([int][char]$ch) {
            8 { [void]$sb.Append('\b') }
            9 { [void]$sb.Append('\t') }
            10 { [void]$sb.Append('\n') }
            12 { [void]$sb.Append('\f') }
            13 { [void]$sb.Append('\r') }
            34 { [void]$sb.Append('\"') }
            92 { [void]$sb.Append('\\') }
            default {
                if ([int][char]$ch -lt 32) {
                    [void]$sb.Append(('\u{0:x4}' -f [int][char]$ch))
                }
                else {
                    [void]$sb.Append($ch)
                }
            }
        }
    }
    return $sb.ToString()
}

function Convert-ToJsonLite {
    param(
        $InputObject,
        [int]$Depth = 0,
        [int]$MaxDepth = 12
    )

    if ($Depth -gt $MaxDepth) { return 'null' }
    if ($null -eq $InputObject) { return 'null' }

    if ($InputObject -is [string]) {
        return ('"' + (Escape-JsonString $InputObject) + '"')
    }

    if ($InputObject -is [bool]) {
        if ($InputObject) { return 'true' } else { return 'false' }
    }

    if ($InputObject -is [DateTime]) {
        return ('"' + $InputObject.ToString('o') + '"')
    }

    if ($InputObject -is [int] -or $InputObject -is [long] -or $InputObject -is [double] -or $InputObject -is [decimal] -or $InputObject -is [single]) {
        return ([string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0}', $InputObject))
    }

    if ($InputObject -is [System.Collections.IDictionary]) {
        $parts = New-List
        foreach ($key in $InputObject.Keys) {
            $jsonKey = '"' + (Escape-JsonString ([string]$key)) + '"'
            $jsonValue = Convert-ToJsonLite -InputObject $InputObject[$key] -Depth ($Depth + 1) -MaxDepth $MaxDepth
            [void]$parts.Add(($jsonKey + ':' + $jsonValue))
        }
        return ('{' + ($parts -join ',') + '}')
    }

    if (($InputObject -is [System.Collections.IEnumerable]) -and -not ($InputObject -is [string])) {
        $parts = New-List
        foreach ($item in $InputObject) {
            [void]$parts.Add((Convert-ToJsonLite -InputObject $item -Depth ($Depth + 1) -MaxDepth $MaxDepth))
        }
        return ('[' + ($parts -join ',') + ']')
    }

    $props = New-Map
    foreach ($prop in $InputObject.PSObject.Properties) {
        $props[$prop.Name] = $prop.Value
    }
    return (Convert-ToJsonLite -InputObject $props -Depth ($Depth + 1) -MaxDepth $MaxDepth)
}

function Write-Utf8File {
    param(
        [string]$Path,
        [string]$Text
    )
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -Path $dir -ItemType Directory -Force | Out-Null
    }
    $utf8 = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $utf8)
}

function Build-TextReport {
    param($Report)
    $lines = New-List
    [void]$lines.Add(($script:ToolName + ' ' + $script:ToolVersion))
    [void]$lines.Add(('Gerado em: ' + $Report.Meta.GeneratedAt))
    [void]$lines.Add(('Host: ' + $Report.Meta.Host))
    [void]$lines.Add(('Usuário: ' + $Report.Meta.ExecutedBy))
    [void]$lines.Add(('Modo: ' + $Report.Meta.Mode))
    [void]$lines.Add(('Role: ' + $Report.Context.Role))
    [void]$lines.Add('')
    [void]$lines.Add('STATUS')
    [void]$lines.Add(('  Status: ' + $Report.Analysis.Status))
    [void]$lines.Add(('  Categoria principal: ' + $Report.Analysis.PrimaryCategory))
    [void]$lines.Add(('  Causa provável: ' + $Report.Analysis.PrimaryCause))
    [void]$lines.Add(('  Confiança: ' + $Report.Analysis.Confidence))
    [void]$lines.Add('')
    [void]$lines.Add('CAMINHOS')
    [void]$lines.Add(('  DesiredDbPath : ' + [string]$Report.Context.DesiredDbPath))
    [void]$lines.Add(('  EffectiveDbPath: ' + [string]$Report.Context.EffectiveDbPath))
    [void]$lines.Add(('  DesiredNetDir : ' + [string]$Report.Context.DesiredNetDir))
    [void]$lines.Add(('  EffectiveNetDir: ' + [string]$Report.Context.EffectiveNetDir))
    [void]$lines.Add(('  ExePath       : ' + [string]$Report.Context.ExpectedExePath))
    [void]$lines.Add('')
    [void]$lines.Add('ACHADOS')
    foreach ($item in @($Report.Analysis.Findings)) {
        [void]$lines.Add(('  - ' + [string]$item))
    }
    if (@($Report.Analysis.Findings).Count -eq 0) {
        [void]$lines.Add('  - Nenhum achado crítico nesta rodada.')
    }
    [void]$lines.Add('')
    [void]$lines.Add('RECOMENDAÇÕES')
    foreach ($item in @($Report.Analysis.Recommendations)) {
        [void]$lines.Add(('  - ' + [string]$item))
    }
    if (@($Report.Analysis.Recommendations).Count -eq 0) {
        [void]$lines.Add('  - Sem recomendação corretiva imediata.')
    }
    [void]$lines.Add('')
    [void]$lines.Add('PLANO DE MUDANÇAS')
    foreach ($item in @($Report.ChangePlan)) {
        if ($null -eq $item) { continue }
        $backupInfo = ''
        try { if (-not [string]::IsNullOrWhiteSpace([string]$item.Backup)) { $backupInfo = ' | Backup=' + [string]$item.Backup } } catch {}
        [void]$lines.Add(('  - ' + [string]$item.Action + ' | ' + [string]$item.Target + ' | ' + [string]$item.Status + ' | ' + [string]$item.Message + $backupInfo))
    }
    if (@(@($Report.ChangePlan) | Where-Object { $null -ne $_ }).Count -eq 0) {
        [void]$lines.Add('  - Nenhuma mudança pendente.')
    }
    [void]$lines.Add('')
    [void]$lines.Add('MUDANÇAS APLICADAS')
    foreach ($item in @($Report.AppliedChanges)) {
        if ($null -eq $item) { continue }
        $backupInfo = ''
        try { if (-not [string]::IsNullOrWhiteSpace([string]$item.Backup)) { $backupInfo = ' | Backup=' + [string]$item.Backup } } catch {}
        [void]$lines.Add(('  - ' + [string]$item.Action + ' | ' + [string]$item.Target + ' | ' + [string]$item.Status + ' | ' + [string]$item.Message + $backupInfo))
    }
    if (@(@($Report.AppliedChanges) | Where-Object { $null -ne $_ }).Count -eq 0) {
        [void]$lines.Add('  - Nenhuma mudança aplicada nesta execução.')
    }
    [void]$lines.Add('')
    [void]$lines.Add('TIMELINE')
    [void]$lines.Add(('  Samples             : ' + [string]$Report.TimelineSummary.SampleCount))
    [void]$lines.Add(('  DbUnavailableSamples: ' + [string]$Report.TimelineSummary.DbUnavailableSamples))
    [void]$lines.Add(('  NetUnavailableSamples: ' + [string]$Report.TimelineSummary.NetDirUnavailableSamples))
    [void]$lines.Add(('  PeakLockCount       : ' + [string]$Report.TimelineSummary.PeakLockCount))
    [void]$lines.Add('')
    [void]$lines.Add('LOG')
    foreach ($line in @($Report.Log)) {
        [void]$lines.Add(('  ' + [string]$line))
    }
    return ($lines -join [Environment]::NewLine)
}

function Build-HtmlReport {
    param($Report)

    $findingsHtml = ''
    foreach ($item in @($Report.Analysis.Findings)) {
        $findingsHtml += ('<li>' + (Escape-Html ([string]$item)) + '</li>')
    }
    if ([string]::IsNullOrWhiteSpace($findingsHtml)) {
        $findingsHtml = '<li>Nenhum achado crítico nesta rodada.</li>'
    }

    $recoHtml = ''
    foreach ($item in @($Report.Analysis.Recommendations)) {
        $recoHtml += ('<li>' + (Escape-Html ([string]$item)) + '</li>')
    }
    if ([string]::IsNullOrWhiteSpace($recoHtml)) {
        $recoHtml = '<li>Sem recomendação corretiva imediata.</li>'
    }

    $planRows = ''
    foreach ($item in @($Report.ChangePlan)) {
        if ($null -eq $item) { continue }
        $planRows += ('<tr><td>' + (Escape-Html ([string]$item.Action)) + '</td><td>' + (Escape-Html ([string]$item.Target)) + '</td><td>' + (Escape-Html ([string]$item.CurrentValue)) + '</td><td>' + (Escape-Html ([string]$item.DesiredValue)) + '</td><td>' + (Escape-Html ([string]$item.Safety)) + '</td></tr>')
    }
    if ([string]::IsNullOrWhiteSpace($planRows)) {
        $planRows = '<tr><td colspan="5">Nenhuma mudança pendente.</td></tr>'
    }

    $appliedRows = ''
    foreach ($item in @($Report.AppliedChanges)) {
        if ($null -eq $item) { continue }
        $appliedRows += ('<tr><td>' + (Escape-Html ([string]$item.Action)) + '</td><td>' + (Escape-Html ([string]$item.Target)) + '</td><td>' + (Escape-Html ([string]$item.Status)) + '</td><td>' + (Escape-Html ([string]$item.Message)) + '</td><td>' + (Escape-Html ([string]$item.Backup)) + '</td></tr>')
    }
    if ([string]::IsNullOrWhiteSpace($appliedRows)) {
        $appliedRows = '<tr><td colspan="5">Nenhuma mudança aplicada nesta execução.</td></tr>'
    }

    $timelineRows = ''
    foreach ($item in @($Report.Timeline)) {
        $timelineRows += ('<tr><td>' + (Escape-Html ([string]$item.Timestamp)) + '</td><td>' + (Escape-Html ([string]$item.DbAccessible)) + '</td><td>' + (Escape-Html ([string]$item.NetDirAccessible)) + '</td><td>' + (Escape-Html ([string]$item.LockCount)) + '</td><td>' + (Escape-Html ([string]$item.Error)) + '</td></tr>')
    }
    if ([string]::IsNullOrWhiteSpace($timelineRows)) {
        $timelineRows = '<tr><td colspan="5">Sem timeline (MonitorMinutes=0).</td></tr>'
    }

    $statusClass = 'ok'
    if ([string]$Report.Analysis.Status -eq 'FALHA') { $statusClass = 'fail' }
    elseif ([string]$Report.Analysis.Status -eq 'ALERTA') { $statusClass = 'warn' }

    $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>ECGv6 FieldKit</title>
<style>
body { font-family: Segoe UI, Tahoma, Arial; margin: 24px; color: #111827; }
.card { border: 1px solid #d1d5db; border-radius: 8px; padding: 16px; margin-bottom: 16px; }
.badge { display: inline-block; padding: 6px 10px; border-radius: 999px; font-weight: 700; color: #fff; }
.ok { background: #059669; }
.warn { background: #d97706; }
.fail { background: #dc2626; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #e5e7eb; padding: 8px; text-align: left; vertical-align: top; }
th { background: #f9fafb; }
ul { margin-top: 8px; }
code { background: #f3f4f6; padding: 2px 4px; }
</style>
</head>
<body>
<h1>ECGv6 FieldKit</h1>
<div class="card">
  <div class="badge $statusClass">$([string](Escape-Html ([string]$Report.Analysis.Status)))</div>
  <p><strong>Causa provável:</strong> $([string](Escape-Html ([string]$Report.Analysis.PrimaryCause)))</p>
  <p><strong>Confiança:</strong> $([string](Escape-Html ([string]$Report.Analysis.Confidence)))</p>
  <p><strong>Host:</strong> $([string](Escape-Html ([string]$Report.Meta.Host))) | <strong>Usuário:</strong> $([string](Escape-Html ([string]$Report.Meta.ExecutedBy))) | <strong>Role:</strong> $([string](Escape-Html ([string]$Report.Context.Role)))</p>
</div>
<div class="card">
  <h2>Caminhos</h2>
  <table>
    <tr><th>Item</th><th>Valor</th></tr>
    <tr><td>DesiredDbPath</td><td><code>$([string](Escape-Html ([string]$Report.Context.DesiredDbPath)))</code></td></tr>
    <tr><td>EffectiveDbPath</td><td><code>$([string](Escape-Html ([string]$Report.Context.EffectiveDbPath)))</code></td></tr>
    <tr><td>DesiredNetDir</td><td><code>$([string](Escape-Html ([string]$Report.Context.DesiredNetDir)))</code></td></tr>
    <tr><td>EffectiveNetDir</td><td><code>$([string](Escape-Html ([string]$Report.Context.EffectiveNetDir)))</code></td></tr>
    <tr><td>ExpectedExePath</td><td><code>$([string](Escape-Html ([string]$Report.Context.ExpectedExePath)))</code></td></tr>
  </table>
</div>
<div class="card">
  <h2>Achados</h2>
  <ul>$findingsHtml</ul>
  <h2>Recomendações</h2>
  <ul>$recoHtml</ul>
</div>
<div class="card">
  <h2>Plano de mudanças</h2>
  <table>
    <tr><th>Ação</th><th>Alvo</th><th>Atual</th><th>Desejado</th><th>Safety</th></tr>
    $planRows
  </table>
</div>
<div class="card">
  <h2>Mudanças aplicadas</h2>
  <table>
    <tr><th>Ação</th><th>Alvo</th><th>Status</th><th>Mensagem</th><th>Backup</th></tr>
    $appliedRows
  </table>
</div>
<div class="card">
  <h2>Timeline</h2>
  <table>
    <tr><th>Timestamp</th><th>DB OK</th><th>NetDir OK</th><th>Locks</th><th>Erro</th></tr>
    $timelineRows
  </table>
</div>
</body>
</html>
"@
    return $html
}

function Open-IfRequested {
    param([string]$Path)
    if (-not $OpenReport) { return }
    try {
        Start-Process -FilePath $Path | Out-Null
    }
    catch {
        Warn ('Não foi possível abrir o relatório automaticamente: ' + $_.Exception.Message)
    }
}

try {
    Log ($script:ToolName + ' ' + $script:ToolVersion + ' iniciado.') 'STEP'

    $profile = Get-ProfileValueMap -Path $ProfilePath
    if (-not [string]::IsNullOrWhiteSpace([string](Get-ProfileOutDir -Profile $profile))) { $OutDir = Get-ProfileOutDir -Profile $profile }
    if ($MonitorMinutes -eq 3 -and -not [string]::IsNullOrWhiteSpace([string](Get-MapValue -Map $profile -Key 'MonitorMinutes' -Default ''))) { $MonitorMinutes = [int](Get-MapValue -Map $profile -Key 'MonitorMinutes' -Default '3') }
    if ($SampleIntervalSeconds -eq 15 -and -not [string]::IsNullOrWhiteSpace([string](Get-MapValue -Map $profile -Key 'SampleIntervalSeconds' -Default ''))) { $SampleIntervalSeconds = [int](Get-MapValue -Map $profile -Key 'SampleIntervalSeconds' -Default '15') }

    if (-not (Test-Path -LiteralPath $OutDir)) {
        New-Item -Path $OutDir -ItemType Directory -Force | Out-Null
    }

    $runId = (Get-Date -Format 'yyyyMMdd_HHmmss') + '_' + $script:HostName
    $runRoot = Join-Path $OutDir $runId
    $backupRoot = Join-Path $runRoot 'backup'
    New-Item -Path $runRoot -ItemType Directory -Force | Out-Null
    New-Item -Path $backupRoot -ItemType Directory -Force | Out-Null
    $envState = Get-UserAndMachineEnvState
    $shares = Get-LocalShares
    $networkConnections = Get-NetworkConnections
    $bdeState = Get-BdeNetDirState
    $heartWareState = Get-HeartWareRegistryState
    $idapiState = Get-IdapiCfgState
    $osState = Get-OperatingSystemState
    $csState = Get-ComputerSystemState
    $identity = Get-CurrentIdentity
    $procInfo = Get-ProcessInventory

    $desiredState = Resolve-DesiredState -Profile $profile -EnvState $envState -HeartWareSnapshots $heartWareState -Shares $shares -NetworkConnections $networkConnections -ExpectedDbPathParam $ExpectedDbPath -ExpectedNetDirParam $ExpectedNetDir -ExpectedExePathParam $ExpectedExePath
    $currentState = Get-CurrentState -EnvState $envState -BdeNetDirState $bdeState -HeartWareSnapshots $heartWareState -Shares $shares -NetworkConnections $networkConnections
    $role = Get-RoleHint -Profile $profile -Shares $shares -DesiredDbPath $desiredState.PreferredDbPath -ExpectedExePath $desiredState.PreferredExePath -OsState $osState


    if ($Mode -eq 'Prepare') {
        $prepared = Prepare-ExpectedPaths -EffectiveOutDir $OutDir -DbPath $desiredState.PreferredDbPath -NetDir $desiredState.PreferredNetDir -ExePath $desiredState.PreferredExePath
        $prepTxt = Join-Path $runRoot 'ECGv6_FieldKit_Prepare.txt'
        Write-Utf8File -Path $prepTxt -Text (Build-PrepareTextReport -Rows $prepared)
        foreach ($row in @($prepared)) { Log (($row.Status + ' | ' + $row.Path + ' | ' + $row.Message)) 'INFO' }
        Log ('Preparo concluído em: ' + $runRoot) 'STEP'
        Open-IfRequested -Path $prepTxt
        exit 0
    }

    if ($Mode -eq 'Compare') {
        $resolvedCompare = Resolve-CompareReportPaths -RootOutDir $OutDir -LeftPath $CompareLeftReport -RightPath $CompareRightReport
        $CompareLeftReport = [string]$resolvedCompare.Left
        $CompareRightReport = [string]$resolvedCompare.Right
        Log ('Compare Left : ' + $CompareLeftReport) 'INFO'
        Log ('Compare Right: ' + $CompareRightReport) 'INFO'
        $leftReport = Read-JsonReportObject -Path $CompareLeftReport
        $rightReport = Read-JsonReportObject -Path $CompareRightReport
        $compare = Build-CompareReport -Left $leftReport -Right $rightReport -LeftPath $CompareLeftReport -RightPath $CompareRightReport
        $cmpJson = Join-Path $runRoot 'ECGv6_FieldKit_Compare.json'
        $cmpTxt = Join-Path $runRoot 'ECGv6_FieldKit_Compare.txt'
        $cmpHtml = Join-Path $runRoot 'ECGv6_FieldKit_Compare.html'
        Write-Utf8File -Path $cmpJson -Text (Convert-ToJsonLite -InputObject $compare)
        Write-Utf8File -Path $cmpTxt -Text (Build-CompareTextReport -Compare $compare)
        Write-Utf8File -Path $cmpHtml -Text (Build-CompareHtmlReport -Compare $compare)
        Log ('Comparação concluída em: ' + $runRoot) 'STEP'
        Open-IfRequested -Path $cmpHtml
        exit 0
    }

    if ($Mode -eq 'Rollback') {
        $ok = Invoke-RollbackReg -Path $RollbackFile
        if ($ok) {
            Log ('Rollback concluído com sucesso: ' + $RollbackFile) 'STEP'
            exit 0
        }
        exit 1
    }

    $dbInspection = Inspect-Path -Path $currentState.EffectiveDbPath -DoAcl:$IncludeAcl
    $netInspection = Inspect-Path -Path $currentState.EffectiveNetDir -DoWriteProbe:$WriteProbe -DoAcl:$IncludeAcl
    $exeInspection = Inspect-Path -Path $desiredState.PreferredExePath
    $writeProbeResult = $null
    if ($netInspection.WriteProbe) { $writeProbeResult = $netInspection.WriteProbe }

    $plan = Build-ChangePlan -DesiredState $desiredState -CurrentState $currentState -EnvState $envState -BdeNetDirState $bdeState
    foreach ($item in @($plan)) { Add-ListItem -List $script:PendingChanges -Value $item }

    $continueExecution = Ensure-AdminForFixes -Plan $plan
    if (-not $continueExecution -and -not (Test-IsAdmin)) {
        exit 0
    }

    $applyResults = New-List
    if ($Mode -eq 'Fix') {
        if ($plan.Count -gt 0) {
            Log ('Aplicando ' + [string]$plan.Count + ' mudança(s) idempotente(s).') 'STEP'
            $applyResults = Apply-ChangePlan -Plan $plan -BackupDir $backupRoot
        }

        if (-not [string]::IsNullOrWhiteSpace($desiredState.PreferredNetDir)) {
            $fixedCfg = Repair-IdapiCfgNetDir -DesiredNetDir $desiredState.PreferredNetDir
            $cfgRow = New-Map
            $cfgRow.Key = 'IDAPI32CFG|' + [string]$fixedCfg.FilePath
            $cfgRow.Action = 'REPAIR_IDAPI_NETDIR'
            $cfgRow.Target = [string]$fixedCfg.FilePath
            $cfgRow.Status = [string]$fixedCfg.Status
            $cfgRow.Message = [string]$fixedCfg.Message
            $cfgRow.Backup = [string]$fixedCfg.BackupPath
            Add-ListItem -List $script:AppliedChanges -Value $cfgRow
            if ($applyResults -is [System.Collections.IList]) { [void]$applyResults.Add($cfgRow) }
            if ($fixedCfg.Status -eq 'APPLIED') {
                Log 'IDAPI32.CFG foi ajustado. Recomenda-se reiniciar o BDE/ECGv6.' -Level 'STEP'
            }
            elseif ($fixedCfg.Status -eq 'WARN') {
                Warn ('Falha ao ajustar IDAPI32.CFG: ' + [string]$fixedCfg.Message)
            }
        }
        $envState = Get-UserAndMachineEnvState
        $bdeState = Get-BdeNetDirState
        $heartWareState = Get-HeartWareRegistryState
        $idapiState = Get-IdapiCfgState
        $currentState = Get-CurrentState -EnvState $envState -BdeNetDirState $bdeState -HeartWareSnapshots $heartWareState -Shares $shares -NetworkConnections $networkConnections
        $dbInspection = Inspect-Path -Path $currentState.EffectiveDbPath -DoAcl:$IncludeAcl
        $netInspection = Inspect-Path -Path $currentState.EffectiveNetDir -DoWriteProbe:$WriteProbe -DoAcl:$IncludeAcl
        if ($netInspection.WriteProbe) { $writeProbeResult = $netInspection.WriteProbe }
    }

    $timeline = Collect-Timeline -DbPath $currentState.EffectiveDbPath -NetDirPath $currentState.EffectiveNetDir -Minutes $MonitorMinutes -IntervalSeconds $SampleIntervalSeconds
    $timelineSummary = Get-TimelineSummary -Timeline $timeline
    $shareHealth = Get-ShareHealthState -Shares $shares -DesiredDbPath $desiredState.PreferredDbPath
    $analysis = Get-ConsistencyState -DesiredState $desiredState -CurrentState $currentState -BdeNetDirState $bdeState -EnvState $envState -DbInspection $dbInspection -NetInspection $netInspection -WriteProbeResult $writeProbeResult -TimelineSummary $timelineSummary -ShareHealth $shareHealth -IdapiCfgState $idapiState -HeartWareSnapshots $heartWareState

    $context = New-Map
    $context.Role = $role
    $context.StationAlias = Get-ProfileAlias -Profile $profile
    $context.DesiredDbPath = $desiredState.PreferredDbPath
    $context.EffectiveDbPath = $currentState.EffectiveDbPath
    $context.DesiredNetDir = $desiredState.PreferredNetDir
    $context.EffectiveNetDir = $currentState.EffectiveNetDir
    $context.ExpectedExePath = $desiredState.PreferredExePath
    $context.SymptomText = $SymptomText
    $context.IsAdmin = Test-IsAdmin

    $meta = New-Map
    $meta.ToolName = $script:ToolName
    $meta.ToolVersion = $script:ToolVersion
    $meta.GeneratedAt = (Get-Date).ToString('o')
    $meta.RunId = $runId
    $meta.Host = $script:HostName
    $meta.ExecutedBy = $identity.Name
    $meta.Mode = $Mode
    $meta.PowerShellVersion = $PSVersionTable.PSVersion.ToString()

    $report = New-Map
    $report.Meta = $meta
    $report.Context = $context
    $report.Environment = New-Map
    $report.Environment.Os = $osState
    $report.Environment.ComputerSystem = $csState
    $report.Environment.Identity = $identity
    $report.Environment.EnvState = $envState
    $report.Environment.BdeNetDirState = $bdeState
    $report.Environment.HeartWareRegistry = $heartWareState
    $report.Environment.IdapiCfg = $idapiState
    $report.Environment.Shares = $shares
    $report.Environment.NetworkConnections = $networkConnections
    $report.Environment.Processes = $procInfo
    $report.Inspections = New-Map
    $report.Inspections.Database = $dbInspection
    $report.Inspections.NetDir = $netInspection
    $report.Inspections.Executable = $exeInspection
    $report.ChangePlan = $plan
    $report.ApplyResults = $applyResults
    $report.AppliedChanges = $script:AppliedChanges
    $report.Timeline = $timeline
    $report.TimelineSummary = $timelineSummary
    $report.Analysis = $analysis
    $report.Warnings = $script:Warnings
    $report.Errors = $script:ErrorsFound
    $report.Log = $script:LogLines

    $jsonPath = Join-Path $runRoot 'ECGv6_FieldKit_Report.json'
    $txtPath = Join-Path $runRoot 'ECGv6_FieldKit_Report.txt'
    $htmlPath = Join-Path $runRoot 'ECGv6_FieldKit_Report.html'
    $logPath = Join-Path $runRoot 'ECGv6_FieldKit_Run.log'

    Write-Utf8File -Path $jsonPath -Text (Convert-ToJsonLite -InputObject $report)
    Write-Utf8File -Path $txtPath -Text (Build-TextReport -Report $report)
    Write-Utf8File -Path $htmlPath -Text (Build-HtmlReport -Report $report)
    Write-Utf8File -Path $logPath -Text (($script:LogLines -join [Environment]::NewLine))

    Log ('Relatórios gravados em: ' + $runRoot) 'STEP'
    Open-IfRequested -Path $htmlPath
}
catch {
    Fail-Note ('Falha fatal: ' + $_.Exception.Message)
    throw
}
