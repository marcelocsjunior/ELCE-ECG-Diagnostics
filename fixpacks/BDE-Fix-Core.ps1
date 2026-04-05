[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('UN1','UN2','UN3','CUSTOM')]
    [string]$Profile = 'UN1',

    [string]$CustomDbPath = '',

    [string]$CustomNetDir = '',

    [ValidateSet('DIAG','NETDIR','HW_CAMINHO_DB','DIRECTORIES','ALL')]
    [string]$TaskMode = 'DIAG',

    [ValidateSet('User','Machine','Process')]
    [string]$HwScope = 'User',

    [ValidateSet('ABRIR_EXAME','SALVAR_FINALIZAR','GERAL')]
    [string]$StagePriority = 'ABRIR_EXAME',

    [ValidateSet('LENTIDAO_TRAVAMENTO','INCONCLUSIVO')]
    [string]$SymptomCode = 'LENTIDAO_TRAVAMENTO',

    [ValidateRange(1,120)]
    [int]$ObservationMinutes = 10,

    [ValidateRange(5,300)]
    [int]$SampleIntervalSeconds = 20,

    [switch]$OpenReportOnSuccess,

    [string]$ProfilesFile = '',

    [string]$DiagnosticScript = '',

    [string]$OutputRoot = 'C:\ECG\Output\Fixes'
)

$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8

$ScriptRoot = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Path }
$RegPathNetDir = 'HKLM:\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT'
$DefaultDirs = @('C:\HW\NetDir', 'C:\HW\Private')
$LatestReportPath = 'C:\ECG\Output\Latest\ELCE_ECG_Diagnostics_Report.html'

function Ensure-Directory {
    param([string]$Path)
    if (-not [string]::IsNullOrWhiteSpace($Path) -and -not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Write-Utf8NoBomFile {
    param(
        [string]$Path,
        [string]$Content
    )
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir)) { Ensure-Directory -Path $dir }
    $enc = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($Path, $Content, $enc)
}

function Save-JsonFile {
    param(
        [string]$Path,
        $Object,
        [int]$Depth = 8
    )
    $json = if ($null -eq $Object) { 'null' } else { $Object | ConvertTo-Json -Depth $Depth }
    Write-Utf8NoBomFile -Path $Path -Content ([string]$json)
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
    param([string]$Path)
    $p = Normalize-PathString $Path
    if ([string]::IsNullOrWhiteSpace($p)) { return '' }
    return $p.ToLowerInvariant()
}

function Convert-ToEnvTarget {
    param([string]$Scope)
    switch ($Scope) {
        'User'    { return [System.EnvironmentVariableTarget]::User }
        'Machine' { return [System.EnvironmentVariableTarget]::Machine }
        'Process' { return [System.EnvironmentVariableTarget]::Process }
        default   { throw "Escopo HW_CAMINHO_DB inválido: $Scope" }
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

function Get-ReinvokeArgumentList {
    $args = New-Object System.Collections.Generic.List[string]
    $args.Add('-NoProfile')
    $args.Add('-ExecutionPolicy')
    $args.Add('Bypass')
    $args.Add('-File')
    $args.Add($MyInvocation.MyCommand.Path)
    $args.Add('-Profile'); $args.Add($Profile)
    if (-not [string]::IsNullOrWhiteSpace($CustomDbPath)) { $args.Add('-CustomDbPath'); $args.Add($CustomDbPath) }
    if (-not [string]::IsNullOrWhiteSpace($CustomNetDir)) { $args.Add('-CustomNetDir'); $args.Add($CustomNetDir) }
    $args.Add('-TaskMode'); $args.Add($TaskMode)
    $args.Add('-HwScope'); $args.Add($HwScope)
    $args.Add('-StagePriority'); $args.Add($StagePriority)
    $args.Add('-SymptomCode'); $args.Add($SymptomCode)
    $args.Add('-ObservationMinutes'); $args.Add([string]$ObservationMinutes)
    $args.Add('-SampleIntervalSeconds'); $args.Add([string]$SampleIntervalSeconds)
    if ($OpenReportOnSuccess) { $args.Add('-OpenReportOnSuccess') }
    if (-not [string]::IsNullOrWhiteSpace($ProfilesFile)) { $args.Add('-ProfilesFile'); $args.Add($ProfilesFile) }
    if (-not [string]::IsNullOrWhiteSpace($DiagnosticScript)) { $args.Add('-DiagnosticScript'); $args.Add($DiagnosticScript) }
    if (-not [string]::IsNullOrWhiteSpace($OutputRoot)) { $args.Add('-OutputRoot'); $args.Add($OutputRoot) }
    if ($WhatIfPreference) { $args.Add('-WhatIf') }
    return $args
}

function Restart-WithBypassIfNeeded {
    if (-not $MyInvocation.MyCommand.Path) { return }
    if ($env:ELCE_UNIFIED_CORE_BYPASS_RESTARTED -eq '1') { return }

    $effectivePolicy = $null
    try { $effectivePolicy = Get-ExecutionPolicy -ErrorAction Stop } catch { $effectivePolicy = $null }

    if ($effectivePolicy -in @('Restricted','AllSigned')) {
        Write-Host "Política de execução '$effectivePolicy' detectada. Reiniciando com Bypass..." -ForegroundColor Yellow
        $args = Get-ReinvokeArgumentList
        $psi = New-Object System.Diagnostics.ProcessStartInfo
        $psi.FileName = 'powershell.exe'
        $psi.UseShellExecute = $false
        $psi.Arguments = ($args | ForEach-Object {
            if ($_ -match '\s') { '"' + ($_ -replace '"', '""') + '"' } else { $_ }
        }) -join ' '
        $psi.EnvironmentVariables['ELCE_UNIFIED_CORE_BYPASS_RESTARTED'] = '1'
        $proc = [System.Diagnostics.Process]::Start($psi)
        $proc.WaitForExit()
        exit $proc.ExitCode
    }
}

function Restart-ElevatedIfNeeded {
    param([bool]$NeedElevation)
    if (-not $NeedElevation) { return }
    if (Test-IsAdmin) { return }

    Write-Host "Elevação necessária para concluir as alterações solicitadas." -ForegroundColor Yellow
    $args = Get-ReinvokeArgumentList
    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName = 'powershell.exe'
    $psi.UseShellExecute = $true
    $psi.Verb = 'runas'
    $psi.Arguments = ($args | ForEach-Object {
        if ($_ -match '\s') { '"' + ($_ -replace '"', '""') + '"' } else { $_ }
    }) -join ' '
    $proc = [System.Diagnostics.Process]::Start($psi)
    $proc.WaitForExit()
    exit $proc.ExitCode
}

function Get-DefaultProfiles {
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
      "fallbackDbPath": "P:\\ECG\\HW\\Database"
    },
    "UN2": {
      "name": "Unidade 2",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN2-ECG\\hw",
      "netDirPath": "\\\\ELCUN2-ECG\\hw\\NetDir",
      "fallbackDbPath": ""
    },
    "UN3": {
      "name": "Unidade 3",
      "topology": "SERVIDOR_LOCAL_EXAME",
      "dbPath": "\\\\ELCUN3-ECG\\hw",
      "netDirPath": "\\\\ELCUN3-ECG\\hw\\NetDir",
      "fallbackDbPath": ""
    }
  }
}
'@
    return ($json | ConvertFrom-Json)
}

function Resolve-ProfilesFile {
    param([string]$Candidate)
    if (-not [string]::IsNullOrWhiteSpace($Candidate) -and (Test-Path -LiteralPath $Candidate)) { return $Candidate }

    $parent = Split-Path $ScriptRoot -Parent
    foreach ($p in @(
        (Join-Path $ScriptRoot 'ECG_UnitProfiles.json'),
        (Join-Path $parent 'fixpacks\ECG_UnitProfiles.json'),
        (Join-Path $parent 'src\ECG_UnitProfiles.json')
    )) {
        if (Test-Path -LiteralPath $p) { return $p }
    }
    return $null
}

function Load-ProfileCatalog {
    param([string]$ProfileFile)
    $catalog = Get-DefaultProfiles
    $resolvedFile = Resolve-ProfilesFile -Candidate $ProfileFile

    if ($resolvedFile) {
        try {
            $raw = Get-Content -LiteralPath $resolvedFile -Raw -Encoding UTF8
            $external = $raw | ConvertFrom-Json
            if ($null -ne $external -and $null -ne $external.units) {
                return [pscustomobject]@{ Catalog = $external; Source = $resolvedFile; Warning = '' }
            }
        }
        catch {
            return [pscustomobject]@{
                Catalog = $catalog
                Source  = 'embedded-defaults'
                Warning = ("Falha ao carregar '" + $resolvedFile + "': " + $_.Exception.Message)
            }
        }
    }

    return [pscustomobject]@{ Catalog = $catalog; Source = 'embedded-defaults'; Warning = '' }
}

function Resolve-ProfileSettings {
    param(
        [string]$SelectedProfile,
        [string]$DbPath,
        [string]$NetDirPath,
        $CatalogWrapper
    )

    if ($SelectedProfile -eq 'CUSTOM') {
        if ([string]::IsNullOrWhiteSpace($DbPath)) {
            throw "Profile CUSTOM exige -CustomDbPath."
        }

        $resolvedDb = Normalize-PathString $DbPath
        $resolvedNet = if ([string]::IsNullOrWhiteSpace($NetDirPath)) {
            Join-Path $resolvedDb 'NetDir'
        } else {
            Normalize-PathString $NetDirPath
        }

        return [pscustomobject]@{
            Name           = 'CUSTOM'
            ProfileSource  = $CatalogWrapper.Source
            Warning        = $CatalogWrapper.Warning
            DatabasePath   = $resolvedDb
            NetDirPath     = $resolvedNet
            FallbackDbPath = ''
        }
    }

    $catalog = $CatalogWrapper.Catalog
    if ($null -eq $catalog.units -or -not ($catalog.units.PSObject.Properties.Name -contains $SelectedProfile)) {
        throw "Perfil '$SelectedProfile' não encontrado no catálogo de profiles."
    }

    $unit = $catalog.units.$SelectedProfile
    $db = Normalize-PathString ([string]$unit.dbPath)
    $net = if (-not [string]::IsNullOrWhiteSpace([string]$unit.netDirPath)) {
        Normalize-PathString ([string]$unit.netDirPath)
    } elseif (-not [string]::IsNullOrWhiteSpace($db)) {
        Join-Path $db 'NetDir'
    } else {
        ''
    }

    return [pscustomobject]@{
        Name           = $SelectedProfile
        ProfileSource  = $CatalogWrapper.Source
        Warning        = $CatalogWrapper.Warning
        DatabasePath   = $db
        NetDirPath     = $net
        FallbackDbPath = Normalize-PathString ([string]$unit.fallbackDbPath)
    }
}

function Resolve-DiagnosticScript {
    param([string]$Candidate)
    if (-not [string]::IsNullOrWhiteSpace($Candidate) -and (Test-Path -LiteralPath $Candidate)) { return $Candidate }

    $parent = Split-Path $ScriptRoot -Parent
    foreach ($p in @(
        (Join-Path $parent 'src\ELCE_ECG_Diagnostics.ps1'),
        (Join-Path $ScriptRoot 'ELCE_ECG_Diagnostics.ps1')
    )) {
        if (Test-Path -LiteralPath $p) { return $p }
    }

    return $null
}

function Get-CurrentState {
    param(
        [string]$RegPath,
        [string]$HwScope,
        $ResolvedProfile
    )

    $netDir = $null
    if (Test-Path -LiteralPath $RegPath) {
        try { $netDir = (Get-ItemProperty -Path $RegPath -Name 'NETDIR' -ErrorAction Stop).NETDIR } catch {}
    }

    $dbUser = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::User)
    $dbMachine = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', [System.EnvironmentVariableTarget]::Machine)
    $dbProcess = $env:HW_CAMINHO_DB
    $dbScoped = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', (Convert-ToEnvTarget -Scope $HwScope))

    $dirStates = @()
    foreach ($dir in $DefaultDirs) {
        $dirStates += [pscustomobject]@{ Path = $dir; Exists = [bool](Test-Path -LiteralPath $dir) }
    }

    return [pscustomobject]@{
        CurrentNetDir           = Normalize-PathString $netDir
        CurrentHwDbScoped       = Normalize-PathString $dbScoped
        CurrentHwDbUser         = Normalize-PathString $dbUser
        CurrentHwDbMachine      = Normalize-PathString $dbMachine
        CurrentHwDbProcess      = Normalize-PathString $dbProcess
        TargetDatabaseExists    = [bool](Test-Path -LiteralPath $ResolvedProfile.DatabasePath)
        TargetNetDirExists      = [bool](Test-Path -LiteralPath $ResolvedProfile.NetDirPath)
        TargetPdoxusrsNetExists = [bool](Test-Path -LiteralPath (Join-Path $ResolvedProfile.NetDirPath 'PDOXUSRS.NET'))
        DirectoryStates         = @($dirStates)
    }
}

function Backup-NetDirRegistry {
    param([string]$OutputPath)
    Ensure-Directory -Path (Split-Path -Path $OutputPath -Parent)
    & reg.exe export "HKLM\SOFTWARE\WOW6432Node\Borland\Database Engine\Settings\SYSTEM\INIT" "$OutputPath" /y | Out-Null
}

function Apply-NetDirFix {
    param(
        [string]$RegPath,
        [string]$TargetNetDir,
        [System.Collections.Generic.List[string]]$Actions
    )

    if ([string]::IsNullOrWhiteSpace($TargetNetDir)) {
        $Actions.Add('NETDIR: alvo não resolvido; nenhuma alteração aplicada.')
        return
    }

    $currentNetDir = $null
    if (Test-Path -LiteralPath $RegPath) {
        try { $currentNetDir = (Get-ItemProperty -Path $RegPath -Name 'NETDIR' -ErrorAction Stop).NETDIR } catch {}
    }

    if ((Get-NormalizedPathForCompare $currentNetDir) -eq (Get-NormalizedPathForCompare $TargetNetDir)) {
        $Actions.Add("NETDIR: já estava correto -> $TargetNetDir")
        return
    }

    if ($PSCmdlet.ShouldProcess($RegPath, "Definir NETDIR para $TargetNetDir")) {
        if (-not (Test-Path -LiteralPath $RegPath)) { New-Item -Path $RegPath -Force | Out-Null }
        New-ItemProperty -Path $RegPath -Name 'NETDIR' -PropertyType String -Value $TargetNetDir -Force | Out-Null
        $Actions.Add("NETDIR: atualizado para $TargetNetDir")
    }
}

function Apply-HwDbFix {
    param(
        [string]$TargetDbPath,
        [string]$Scope,
        [System.Collections.Generic.List[string]]$Actions
    )

    if ([string]::IsNullOrWhiteSpace($TargetDbPath)) {
        $Actions.Add('HW_CAMINHO_DB: alvo não resolvido; nenhuma alteração aplicada.')
        return
    }

    $envTarget = Convert-ToEnvTarget -Scope $Scope
    $current = [Environment]::GetEnvironmentVariable('HW_CAMINHO_DB', $envTarget)

    if ((Get-NormalizedPathForCompare $current) -eq (Get-NormalizedPathForCompare $TargetDbPath)) {
        $Actions.Add("HW_CAMINHO_DB($Scope): já estava correto -> $TargetDbPath")
        return
    }

    if ($PSCmdlet.ShouldProcess("HW_CAMINHO_DB($Scope)", "Definir para $TargetDbPath")) {
        [Environment]::SetEnvironmentVariable('HW_CAMINHO_DB', $TargetDbPath, $envTarget)
        if ($Scope -eq 'Process') { $env:HW_CAMINHO_DB = $TargetDbPath }
        $Actions.Add("HW_CAMINHO_DB($Scope): atualizado para $TargetDbPath")
    }
}

function Apply-DirectoriesFix {
    param([System.Collections.Generic.List[string]]$Actions)
    foreach ($dir in $DefaultDirs) {
        if (Test-Path -LiteralPath $dir) {
            $Actions.Add("DIRECTORY: já existe -> $dir")
            continue
        }
        if ($PSCmdlet.ShouldProcess($dir, 'Criar diretório padrão')) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
            $Actions.Add("DIRECTORY: criado -> $dir")
        }
    }
}

function Invoke-DiagnosticMode {
    param(
        [string]$DiagScriptPath,
        [System.Collections.Generic.List[string]]$Actions
    )

    if ([string]::IsNullOrWhiteSpace($DiagScriptPath) -or -not (Test-Path -LiteralPath $DiagScriptPath)) {
        throw "Script de diagnóstico não encontrado. Esperado em '..\src\ELCE_ECG_Diagnostics.ps1' ou caminho informado por -DiagnosticScript."
    }

    $Actions.Add("DIAG: script localizado -> $DiagScriptPath")

    if ($WhatIfPreference) {
        $Actions.Add("DIAG: WhatIf ativo; o diagnóstico não foi executado.")
        return [pscustomobject]@{
            DiagnosticScript = $DiagScriptPath
            LatestReportPath = $LatestReportPath
            ReportExists     = [bool](Test-Path -LiteralPath $LatestReportPath)
            Executed         = $false
        }
    }

    & $DiagScriptPath -StagePriority $StagePriority -SymptomCode $SymptomCode -ObservationMinutes $ObservationMinutes -SampleIntervalSeconds $SampleIntervalSeconds -OpenReportOnSuccess:$OpenReportOnSuccess
    $exitCode = $LASTEXITCODE
    if ($null -eq $exitCode) { $exitCode = 0 }
    if ([int]$exitCode -ne 0) {
        throw "Diagnóstico retornou ExitCode $exitCode."
    }

    $reportExists = [bool](Test-Path -LiteralPath $LatestReportPath)
    if ($reportExists) {
        $Actions.Add("DIAG: relatório HTML gerado em $LatestReportPath")
    } else {
        $Actions.Add("DIAG: execução concluída, mas o HTML não foi encontrado em $LatestReportPath")
    }

    return [pscustomobject]@{
        DiagnosticScript = $DiagScriptPath
        LatestReportPath = $LatestReportPath
        ReportExists     = $reportExists
        Executed         = $true
    }
}

Restart-WithBypassIfNeeded

$needElevation = ($TaskMode -eq 'NETDIR') -or ($TaskMode -eq 'ALL') -or ($TaskMode -eq 'DIRECTORIES') -or (($TaskMode -eq 'HW_CAMINHO_DB' -or $TaskMode -eq 'ALL') -and ($HwScope -eq 'Machine'))
Restart-ElevatedIfNeeded -NeedElevation:$needElevation

$catalogWrapper = Load-ProfileCatalog -ProfileFile $ProfilesFile
$resolvedProfile = Resolve-ProfileSettings -SelectedProfile $Profile -DbPath $CustomDbPath -NetDirPath $CustomNetDir -CatalogWrapper $catalogWrapper
$diagScriptPath = Resolve-DiagnosticScript -Candidate $DiagnosticScript

Ensure-Directory -Path $OutputRoot
$stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$jsonOut = Join-Path $OutputRoot ("Unified_Core_{0}_{1}.json" -f $env:COMPUTERNAME, $stamp)
$txtOut  = Join-Path $OutputRoot ("Unified_Core_{0}_{1}.txt" -f $env:COMPUTERNAME, $stamp)
$backupFile = Join-Path $OutputRoot ("BDE_NETDIR_Backup_{0}_{1}.reg" -f $env:COMPUTERNAME, $stamp)

$before = Get-CurrentState -RegPath $RegPathNetDir -HwScope $HwScope -ResolvedProfile $resolvedProfile
$actions = New-Object System.Collections.Generic.List[string]
$diagResult = $null

Write-Host "=== ELCE ECG Unified Core ===" -ForegroundColor Cyan
Write-Host ""
Write-Host ("TaskMode            : {0}" -f $TaskMode)
Write-Host ("Perfil              : {0}" -f $resolvedProfile.Name)
Write-Host ("Origem do profile   : {0}" -f $resolvedProfile.ProfileSource)
if (-not [string]::IsNullOrWhiteSpace($resolvedProfile.Warning)) {
    Write-Host ("Aviso profile       : {0}" -f $resolvedProfile.Warning) -ForegroundColor Yellow
}
Write-Host ("DB alvo             : {0}" -f $resolvedProfile.DatabasePath)
Write-Host ("NETDIR alvo         : {0}" -f $resolvedProfile.NetDirPath)
Write-Host ("Escopo HW DB        : {0}" -f $HwScope)
Write-Host ""

try {
    switch ($TaskMode) {
        'DIAG' {
            $diagResult = Invoke-DiagnosticMode -DiagScriptPath $diagScriptPath -Actions $actions
        }
        'NETDIR' {
            try { Backup-NetDirRegistry -OutputPath $backupFile; $actions.Add("BACKUP: registro exportado para $backupFile") } catch { $actions.Add("BACKUP: falhou -> $($_.Exception.Message)") }
            Apply-NetDirFix -RegPath $RegPathNetDir -TargetNetDir $resolvedProfile.NetDirPath -Actions $actions
        }
        'HW_CAMINHO_DB' {
            Apply-HwDbFix -TargetDbPath $resolvedProfile.DatabasePath -Scope $HwScope -Actions $actions
        }
        'DIRECTORIES' {
            Apply-DirectoriesFix -Actions $actions
        }
        'ALL' {
            try { Backup-NetDirRegistry -OutputPath $backupFile; $actions.Add("BACKUP: registro exportado para $backupFile") } catch { $actions.Add("BACKUP: falhou -> $($_.Exception.Message)") }
            Apply-NetDirFix -RegPath $RegPathNetDir -TargetNetDir $resolvedProfile.NetDirPath -Actions $actions
            Apply-HwDbFix -TargetDbPath $resolvedProfile.DatabasePath -Scope $HwScope -Actions $actions
            Apply-DirectoriesFix -Actions $actions
        }
    }
}
catch {
    $actions.Add("ERRO: $($_.Exception.Message)")
    $after = Get-CurrentState -RegPath $RegPathNetDir -HwScope $HwScope -ResolvedProfile $resolvedProfile
    $result = [pscustomobject]@{
        GeneratedAt        = (Get-Date).ToString('o')
        ComputerName       = $env:COMPUTERNAME
        TaskMode           = $TaskMode
        Profile            = $resolvedProfile.Name
        ProfileSource      = $resolvedProfile.ProfileSource
        ProfileWarning     = $resolvedProfile.Warning
        DatabasePathTarget = $resolvedProfile.DatabasePath
        NetDirTarget       = $resolvedProfile.NetDirPath
        HwScope            = $HwScope
        Before             = $before
        After              = $after
        DiagnosticResult   = $diagResult
        Actions            = @($actions)
        OutputJson         = $jsonOut
        OutputTxt          = $txtOut
        BackupFile         = $backupFile
        WhatIf             = [bool]$WhatIfPreference
        Success            = $false
        Error              = $_.Exception.Message
    }
    Save-JsonFile -Path $jsonOut -Object $result -Depth 8
    Write-Utf8NoBomFile -Path $txtOut -Content (($result | ConvertTo-Json -Depth 8))
    Write-Host ('JSON: ' + $jsonOut) -ForegroundColor Green
    Write-Host ('TXT : ' + $txtOut) -ForegroundColor Green
    Write-Error $_.Exception.Message
    exit 1
}

$after = Get-CurrentState -RegPath $RegPathNetDir -HwScope $HwScope -ResolvedProfile $resolvedProfile

$result = [pscustomobject]@{
    GeneratedAt        = (Get-Date).ToString('o')
    ComputerName       = $env:COMPUTERNAME
    TaskMode           = $TaskMode
    Profile            = $resolvedProfile.Name
    ProfileSource      = $resolvedProfile.ProfileSource
    ProfileWarning     = $resolvedProfile.Warning
    DatabasePathTarget = $resolvedProfile.DatabasePath
    NetDirTarget       = $resolvedProfile.NetDirPath
    HwScope            = $HwScope
    Before             = $before
    After              = $after
    DiagnosticResult   = $diagResult
    Actions            = @($actions)
    OutputJson         = $jsonOut
    OutputTxt          = $txtOut
    BackupFile         = $backupFile
    WhatIf             = [bool]$WhatIfPreference
    Success            = $true
}

Save-JsonFile -Path $jsonOut -Object $result -Depth 8

$txt = @()
$txt += '============================================'
$txt += 'ELCE ECG UNIFIED CORE - RESULTADO'
$txt += '============================================'
$txt += ('Computador          : {0}' -f $env:COMPUTERNAME)
$txt += ('TaskMode            : {0}' -f $TaskMode)
$txt += ('Perfil              : {0}' -f $resolvedProfile.Name)
$txt += ('Origem do profile   : {0}' -f $resolvedProfile.ProfileSource)
$txt += ('Aviso profile       : {0}' -f $resolvedProfile.Warning)
$txt += ('DB alvo             : {0}' -f $resolvedProfile.DatabasePath)
$txt += ('NETDIR alvo         : {0}' -f $resolvedProfile.NetDirPath)
$txt += ('Escopo HW DB        : {0}' -f $HwScope)
$txt += ('WhatIf              : {0}' -f [string]([bool]$WhatIfPreference))
$txt += ''
$txt += 'ACOES'
$txt += '-----'
if ($actions.Count -gt 0) { foreach ($action in $actions) { $txt += ('- ' + $action) } } else { $txt += '- Nenhuma ação executada.' }
if ($diagResult) {
    $txt += ''
    $txt += 'DIAGNOSTICO'
    $txt += '-----------'
    $txt += ('Script             : {0}' -f $diagResult.DiagnosticScript)
    $txt += ('Relatorio HTML     : {0}' -f $diagResult.LatestReportPath)
    $txt += ('Relatorio existe   : {0}' -f [string]$diagResult.ReportExists)
}
$txt += ''
$txt += ('JSON                : {0}' -f $jsonOut)
$txt += ('TXT                 : {0}' -f $txtOut)
$txt += ('BACKUP REG          : {0}' -f $backupFile)
$txt += '============================================'

Write-Utf8NoBomFile -Path $txtOut -Content ($txt -join [Environment]::NewLine)

Write-Host ('JSON: ' + $jsonOut) -ForegroundColor Green
Write-Host ('TXT : ' + $txtOut) -ForegroundColor Green
if ($diagResult -and $diagResult.ReportExists) {
    Write-Host ('HTML: ' + $diagResult.LatestReportPath) -ForegroundColor Green
}

if ($WhatIfPreference) {
    Write-Host 'Execução concluída em modo WhatIf.' -ForegroundColor Yellow
} else {
    Write-Host 'Execução concluída.' -ForegroundColor Green
}
exit 0
