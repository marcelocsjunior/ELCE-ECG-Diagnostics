<#
.SYNOPSIS
    ECG Diagnostics AI Prototype - camada off-repo para explicação assistida dos laudos JSON.
.DESCRIPTION
    Este protótipo NÃO altera o core nem o hub existentes.
    Opera sobre os JSONs gerados pelo pacote ECG Diagnostics Core v6.3.x.

    Modos:
      Explain    : explica o último laudo ou um laudo informado
      Executive  : gera resumo executivo curto
      Technical  : gera parecer técnico por público
      CompareAI  : compara dois laudos com narrativa
      Ask        : responde pergunta sobre um laudo

    Provider default:
      LocalRules : resposta determinística local, sem API externa

    Providers opcionais:
      OpenAICompatible : endpoint compatível com /chat/completions
#>

[CmdletBinding()]
param(
    [ValidateSet('Explain','Executive','Technical','CompareAI','Ask')]
    [string]$Mode = 'Explain',

    [string]$ProfilePath = '',
    [string]$ConfigPath = '',
    [string]$PromptCatalogPath = '',
    [string]$ReportPath = '',
    [string]$LeftReportPath = '',
    [string]$RightReportPath = '',
    [ValidateSet('DEV','INFRA','CAMPO','EXEC')]
    [string]$Audience = 'INFRA',
    [string]$Question = '',
    [string]$OutDir = '',
    [string]$Provider = '',
    [int]$TimeoutSeconds = 0,
    [switch]$OpenOutput,
    [switch]$NoRedaction,
    [switch]$SavePromptCopy,
    [switch]$SaveRawResponse
)

$ErrorActionPreference = 'Stop'
$script:ToolName = 'ECG Diagnostics AI Prototype'
$script:ToolVersion = '0.2.1-localrules'
$script:ScriptPath = $MyInvocation.MyCommand.Path
$script:ScriptDir = Split-Path -Parent $script:ScriptPath
$script:LogLines = New-Object System.Collections.ArrayList
$script:RunStamp = Get-Date -Format 'yyyyMMdd_HHmmss'
$script:CurrentOutDir = ''
$script:CurrentAiDir = ''

function Log {
    param([string]$Message, [string]$Level = 'INFO')
    $line = "{0} [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    [void]$script:LogLines.Add($line)
    Write-Host $line
}

function Write-Utf8File {
    param([string]$Path, [string]$Text)
    $dir = Split-Path -Path $Path -Parent
    if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText($Path, $Text, $enc)
}

function Ensure-Directory {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return }
    if (-not (Test-Path -LiteralPath $Path)) {
        New-Item -ItemType Directory -Path $Path -Force | Out-Null
    }
}

function Normalize-PathString {
    param([string]$Path)
    if ([string]::IsNullOrWhiteSpace($Path)) { return $null }
    $p = ($Path -replace '/', '\').Trim()
    while ($p.Length -gt 3 -and $p.EndsWith('\')) { $p = $p.Substring(0, $p.Length - 1) }
    return $p
}

function Read-IniLoose {
    param([string]$Path)
    $ini = @{}
    if (-not (Test-Path -LiteralPath $Path)) { return $ini }
    $section = '_flat'
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

function Get-IniValue {
    param($Ini, [string]$Section, [string]$Key, [string]$Default = '')
    if ($Ini.ContainsKey($Section) -and $Ini[$Section].ContainsKey($Key)) {
        $value = [string]$Ini[$Section][$Key]
        if (-not [string]::IsNullOrWhiteSpace($value)) { return $value.Trim() }
    }
    return $Default
}

function Resolve-ProfilePath {
    if (-not [string]::IsNullOrWhiteSpace($ProfilePath) -and (Test-Path -LiteralPath $ProfilePath)) {
        return $ProfilePath
    }
    $candidates = @(
        (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_3_2.ini'),
        (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_3_1.ini'),
        (Join-Path $script:ScriptDir 'ECG_FieldKit_Unified_v6_2.ini'),
        (Join-Path $script:ScriptDir 'ECG_FieldKit.ini'),
        'C:\ECG\FieldKit\ECG_FieldKit_Unified_v6_3_2.ini',
        'C:\ECG\FieldKit\ECG_FieldKit_Unified_v6_3_1.ini',
        'C:\ECG\FieldKit\ECG_FieldKit_Unified_v6_2.ini',
        'C:\ECG\FieldKit\ECG_FieldKit.ini'
    )
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return $c }
    }
    return ''
}

function Resolve-EffectiveOutDir {
    param([string]$ResolvedProfilePath)
    if (-not [string]::IsNullOrWhiteSpace($OutDir)) { return (Normalize-PathString $OutDir) }
    if (-not [string]::IsNullOrWhiteSpace($ResolvedProfilePath) -and (Test-Path -LiteralPath $ResolvedProfilePath)) {
        $ini = Read-IniLoose -Path $ResolvedProfilePath
        $generalOut = Get-IniValue -Ini $ini -Section 'General' -Key 'OutDir' -Default ''
        if (-not [string]::IsNullOrWhiteSpace($generalOut)) { return (Normalize-PathString $generalOut) }
        $flatOut = Get-IniValue -Ini $ini -Section '_flat' -Key 'OutDir' -Default ''
        if (-not [string]::IsNullOrWhiteSpace($flatOut)) { return (Normalize-PathString $flatOut) }
    }
    return 'C:\ECG\FieldKit\out'
}

function Resolve-ConfigPath {
    if (-not [string]::IsNullOrWhiteSpace($ConfigPath) -and (Test-Path -LiteralPath $ConfigPath)) { return $ConfigPath }
    $candidate = Join-Path $script:ScriptDir 'ECG_AI.config.json'
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    return $candidate
}

function Resolve-PromptCatalogPath {
    if (-not [string]::IsNullOrWhiteSpace($PromptCatalogPath) -and (Test-Path -LiteralPath $PromptCatalogPath)) { return $PromptCatalogPath }
    $candidate = Join-Path $script:ScriptDir 'ECG_AI_Prompts.json'
    if (Test-Path -LiteralPath $candidate) { return $candidate }
    return $candidate
}

function Read-JsonFile {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { throw "Arquivo JSON não encontrado: $Path" }
    return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json)
}

function Read-JsonConfigOrDefault {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return [PSCustomObject]@{
            Enabled = $true
            Provider = 'LocalRules'
            Endpoint = ''
            Model = ''
            ApiKeyEnvVar = 'OPENAI_API_KEY'
            TimeoutSeconds = 45
            MaxTokens = 700
            Temperature = 0.2
            RedactSensitiveData = $true
            FallbackWithoutAI = $true
            SavePromptCopy = $false
            SaveRawResponse = $false
            OutputFormat = 'txt'
        }
    }
    return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json)
}

function Read-PromptCatalogOrDefault {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) {
        return [PSCustomObject]@{
            Explain = 'Explique o laudo em linguagem operacional objetiva. Use apenas os fatos fornecidos.'
            Executive = 'Resuma o laudo em linguagem executiva, curta e objetiva. Use apenas os fatos fornecidos.'
            Technical_DEV = 'Escreva parecer técnico para desenvolvedor, priorizando aplicação, comportamento e hipótese técnica.'
            Technical_INFRA = 'Escreva parecer técnico para infraestrutura, priorizando share, SMB, latência, storage, rede e sistema.'
            Technical_CAMPO = 'Escreva orientação operacional para campo, objetiva e executável.'
            Technical_EXEC = 'Escreva parecer executivo curto, com foco em impacto e próxima ação.'
            CompareAI = 'Compare os dois laudos. Aponte melhor alvo, principais diferenças e conclusão.'
            Ask = 'Responda à pergunta do operador com base estrita no laudo.'
        }
    }
    return (Get-Content -LiteralPath $Path -Raw -ErrorAction Stop | ConvertFrom-Json)
}


function Resolve-ReportInputPath {
    param([string]$InputPath)

    if ([string]::IsNullOrWhiteSpace($InputPath)) { return '' }

    $candidate = Normalize-PathString $InputPath
    if (-not (Test-Path -LiteralPath $candidate)) {
        throw "Caminho informado nao encontrado: $candidate"
    }

    $item = Get-Item -LiteralPath $candidate -ErrorAction Stop
    if ($item.PSIsContainer) {
        $patterns = @('ECG_Report.json','Single_Report.json','CompareBackend_Report.json','CompareJson_Report.json')
        foreach ($pattern in $patterns) {
            $match = Get-ChildItem -LiteralPath $candidate -Recurse -File -Filter $pattern -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1
            if ($null -ne $match) { return $match.FullName }
        }
        throw "Nenhum laudo JSON reconhecido foi encontrado dentro da pasta: $candidate"
    }

    if ($item.Extension -ine '.json') {
        throw "O caminho informado nao eh um arquivo JSON: $candidate"
    }

    return $item.FullName
}

function Find-LatestReportPath {
    param([string]$BaseOutDir)

    if (-not (Test-Path -LiteralPath $BaseOutDir)) {
        throw "OutDir não encontrado: $BaseOutDir"
    }

    $patterns = @('ECG_Report.json','Single_Report.json','CompareBackend_Report.json','CompareJson_Report.json')
    $files = @()
    foreach ($pattern in $patterns) {
        $files += @(Get-ChildItem -Path $BaseOutDir -Recurse -Filter $pattern -File -ErrorAction SilentlyContinue)
    }
    $files = @($files | Sort-Object LastWriteTime -Descending)
    if ($files.Count -eq 0) {
        throw "Nenhum laudo JSON encontrado em $BaseOutDir"
    }
    return $files[0].FullName
}

function Get-ReportKind {
    param($Data)
    if ($null -ne $Data.TargetReport) { return 'Single' }
    if ($null -ne $Data.Comparison -and $null -ne $Data.Targets) { return 'CompareBackend' }
    if ($null -ne $Data.Analysis -and $null -ne $Data.PassiveBenchmark) { return 'Core' }
    if ($Data -is [System.Collections.IEnumerable] -and -not ($Data -is [string])) { return 'CompareJsonRows' }
    return 'Unknown'
}

function Mask-Text {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $Text }

    $t = $Text
    $t = [regex]::Replace($t, '\\\\[^\\\s]+\\[^ \r\n\t]+', '\\\\<REDACTED_HOST>\\<REDACTED_SHARE>')
    $t = [regex]::Replace($t, '\b(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)\b', '<REDACTED_IP>')
    $t = [regex]::Replace($t, '(?i)\b[A-Z]:\\[^ \r\n\t]+', '<REDACTED_PATH>')
    return $t
}

function ConvertTo-RedactedObject {
    param($InputObject)

    if ($null -eq $InputObject) { return $null }

    if ($InputObject -is [string]) {
        return (Mask-Text -Text $InputObject)
    }

    if ($InputObject -is [bool] -or $InputObject -is [int] -or $InputObject -is [long] -or $InputObject -is [double] -or $InputObject -is [decimal]) {
        return $InputObject
    }

    if ($InputObject -is [System.Collections.IEnumerable] -and -not ($InputObject -is [string])) {
        $arr = @()
        foreach ($item in @($InputObject)) {
            $arr += ,(ConvertTo-RedactedObject -InputObject $item)
        }
        return ,$arr
    }

    $obj = [ordered]@{}
    foreach ($prop in $InputObject.PSObject.Properties) {
        $name = [string]$prop.Name
        $value = $prop.Value

        if ($name -in @('ComputerName','UserName','ExecutedBy','DbHost','NetDirHost','Host','ServerName')) {
            $obj[$name] = '<REDACTED>'
        }
        else {
            $obj[$name] = ConvertTo-RedactedObject -InputObject $value
        }
    }
    return [PSCustomObject]$obj
}

function Get-TopItems {
    param([object[]]$Items, [int]$Max = 3)
    $list = @()
    foreach ($i in @($Items)) {
        $text = [string]$i
        if (-not [string]::IsNullOrWhiteSpace($text)) { $list += $text.Trim() }
    }
    if ($list.Count -le $Max) { return @($list) }
    return @($list[0..($Max - 1)])
}

function Build-NormalizedSnapshot {
    param($Data, [string]$SourcePath)

    $kind = Get-ReportKind -Data $Data

    switch ($kind) {
        'Core' {
            $metrics = $Data.Analysis.Metrics
            return [PSCustomObject]@{
                Kind = 'Core'
                SourcePath = $SourcePath
                Label = [string]$Data.Metadata.Mode
                Status = [string]$Data.Analysis.Status
                Hypothesis = [string]$Data.Analysis.PrimaryHypothesis
                Confidence = [string]$Data.Analysis.Confidence
                Recommendation = [string]$Data.Analysis.RecommendedAction
                Evidence = @(Get-TopItems -Items $Data.Analysis.HypothesisSupport -Max 4)
                CounterEvidence = @(Get-TopItems -Items $Data.Analysis.WhatDidNotIndicateFailure -Max 4)
                Metrics = [ordered]@{
                    SeverityScore = $metrics.SeverityScore
                    PressureLabel = $metrics.PressureLabel
                    AverageCpuPercent = $metrics.AverageCpuPercent
                    PeakCpuPercent = $metrics.PeakCpuPercent
                    AverageDatabaseProbeMs = $metrics.AverageDatabaseProbeMs
                    P95DatabaseProbeMs = $metrics.P95DatabaseProbeMs
                    AverageNetDirProbeMs = $metrics.AverageNetDirProbeMs
                    P95NetDirProbeMs = $metrics.P95NetDirProbeMs
                    PeakLockFileCount = $metrics.PeakLockFileCount
                    DatabaseUnavailableSamples = $metrics.DatabaseUnavailableSamples
                    NetDirUnavailableSamples = $metrics.NetDirUnavailableSamples
                    SmbTimeoutSamples = $metrics.SmbTimeoutSamples
                    PeakCpuResponsibleSummary = $metrics.PeakCpuResponsibleSummary
                }
                ComparableValue = $metrics.SeverityScore
                ComparableMode = 'LowerIsBetter'
            }
        }
        'Single' {
            $tr = $Data.TargetReport
            return [PSCustomObject]@{
                Kind = 'Single'
                SourcePath = $SourcePath
                Label = [string]$tr.Target.Label
                Status = [string]$tr.Assessment.Classification
                Hypothesis = [string]$tr.Assessment.Classification
                Confidence = [string]$tr.Assessment.Confidence
                Recommendation = (@($tr.Assessment.Actions) -join ' | ')
                Evidence = @(Get-TopItems -Items $tr.Assessment.Findings -Max 4)
                CounterEvidence = @()
                Metrics = [ordered]@{
                    CompositeAvgMs = $tr.Stats.CompositeAvgMs
                    CompositeP95Ms = $tr.Stats.CompositeP95Ms
                    DbAvgMs = $tr.Stats.Db.Avg
                    DbP95Ms = $tr.Stats.Db.P95
                    NetDirAvgMs = $tr.Stats.NetDir.Avg
                    NetDirP95Ms = $tr.Stats.NetDir.P95
                    LocalAvgMs = $tr.Stats.Local.Avg
                    LocalP95Ms = $tr.Stats.Local.P95
                    CpuAvgPercent = $tr.Stats.Cpu.Avg
                    CpuPeakPercent = $tr.Stats.Cpu.Max
                    DbUnavailableSamples = $tr.Stats.DbUnavailableSamples
                    NetDirUnavailableSamples = $tr.Stats.NetUnavailableSamples
                    RemoteToLocalRatio = $tr.Assessment.RemoteToLocalRatio
                    ObservedSmbDialects = (@($tr.Target.ObservedSmbDialects) -join ', ')
                }
                ComparableValue = $tr.Stats.CompositeAvgMs
                ComparableMode = 'LowerIsBetter'
            }
        }
        'CompareBackend' {
            return [PSCustomObject]@{
                Kind = 'CompareBackend'
                SourcePath = $SourcePath
                Label = 'CompareBackend'
                Status = [string]$Data.Comparison.LegacyCorrelation
                Hypothesis = [string]$Data.Comparison.LegacyCorrelation
                Confidence = $(switch ([string]$Data.Comparison.LegacyCorrelation) {
                    'FORTE' { 'Alta' }
                    'MÉDIA' { 'Média' }
                    'FRACA' { 'Média' }
                    'CONTRÁRIA' { 'Alta' }
                    default { 'Baixa' }
                })
                Recommendation = [string]$Data.Comparison.Recommendation
                Evidence = @(
                    ("Melhor alvo: " + [string]$Data.Comparison.BestTargetLabel),
                    [string]$Data.Comparison.Summary
                )
                CounterEvidence = @()
                Metrics = [ordered]@{
                    BestTargetLabel = $Data.Comparison.BestTargetLabel
                    LegacyCorrelation = $Data.Comparison.LegacyCorrelation
                    DeltaMsModernVsLegacy = $Data.Comparison.DeltaMsModernVsLegacy
                    RatioModernVsLegacy = $Data.Comparison.RatioModernVsLegacy
                }
                ComparableValue = $Data.Comparison.DeltaMsModernVsLegacy
                ComparableMode = 'AbsoluteCloserToZeroIsBetter'
            }
        }
        default {
            return [PSCustomObject]@{
                Kind = 'Unknown'
                SourcePath = $SourcePath
                Label = 'Desconhecido'
                Status = 'INCONCLUSIVO'
                Hypothesis = 'Formato não reconhecido'
                Confidence = 'Baixa'
                Recommendation = 'Validar o JSON de entrada.'
                Evidence = @('Formato de laudo não reconhecido pela camada IA.')
                CounterEvidence = @()
                Metrics = [ordered]@{}
                ComparableValue = $null
                ComparableMode = 'None'
            }
        }
    }
}

function Convert-MetricsToLines {
    param($Snapshot)
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($entry in $Snapshot.Metrics.GetEnumerator()) {
        if ($null -ne $entry.Value -and -not [string]::IsNullOrWhiteSpace([string]$entry.Value)) {
            [void]$lines.Add(("{0}: {1}" -f $entry.Key, [string]$entry.Value))
        }
    }
    return @($lines)
}

function Build-LocalExplainText {
    param($Snapshot)
    $evidence = @(Get-TopItems -Items $Snapshot.Evidence -Max 4)
    $counter = @(Get-TopItems -Items $Snapshot.CounterEvidence -Max 4)
    $metricLines = @(Convert-MetricsToLines -Snapshot $Snapshot)

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Situação atual:")
    [void]$sb.AppendLine(("Status: {0}" -f $Snapshot.Status))
    [void]$sb.AppendLine(("Hipótese principal: {0}" -f $Snapshot.Hypothesis))
    [void]$sb.AppendLine(("Confiança: {0}" -f $Snapshot.Confidence))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Evidências que sustentam:")
    if ($evidence.Count -eq 0) { [void]$sb.AppendLine('- Sem evidência dominante adicional.') }
    foreach ($line in $evidence) { [void]$sb.AppendLine('- ' + $line) }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("O que não sustentou falha:")
    if ($counter.Count -eq 0) { [void]$sb.AppendLine('- Não há contraponto relevante registrado para este formato de laudo.') }
    foreach ($line in $counter) { [void]$sb.AppendLine('- ' + $line) }
    if ($metricLines.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine("Métricas-chave:")
        foreach ($line in $metricLines) { [void]$sb.AppendLine('- ' + $line) }
    }
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine("Próximo passo recomendado:")
    [void]$sb.AppendLine($Snapshot.Recommendation)
    return $sb.ToString().TrimEnd()
}

function Build-LocalExecutiveText {
    param($Snapshot)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine("Resumo executivo:")
    [void]$sb.AppendLine(("A rodada foi classificada como {0}, com leitura principal em {1}." -f $Snapshot.Status, $Snapshot.Hypothesis))
    if ($Snapshot.Evidence.Count -gt 0) {
        [void]$sb.AppendLine(("Principal evidência: {0}" -f [string]$Snapshot.Evidence[0]))
    }
    [void]$sb.AppendLine(("Confiança: {0}" -f $Snapshot.Confidence))
    [void]$sb.AppendLine(("Ação recomendada: {0}" -f $Snapshot.Recommendation))
    return $sb.ToString().TrimEnd()
}

function Build-LocalTechnicalText {
    param($Snapshot, [string]$Audience)
    $title = switch ($Audience) {
        'DEV' { 'Parecer técnico para desenvolvedor' }
        'CAMPO' { 'Parecer operacional para campo' }
        'EXEC' { 'Parecer executivo curto' }
        default { 'Parecer técnico para infraestrutura' }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine($title + ':')
    [void]$sb.AppendLine(("Status: {0}" -f $Snapshot.Status))
    [void]$sb.AppendLine(("Hipótese principal: {0}" -f $Snapshot.Hypothesis))
    [void]$sb.AppendLine(("Confiança: {0}" -f $Snapshot.Confidence))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Achados:')
    foreach ($line in @(Get-TopItems -Items $Snapshot.Evidence -Max 5)) { [void]$sb.AppendLine('- ' + $line) }

    $metricLines = @(Convert-MetricsToLines -Snapshot $Snapshot)
    if ($metricLines.Count -gt 0) {
        [void]$sb.AppendLine('')
        [void]$sb.AppendLine('Métricas relevantes:')
        foreach ($line in $metricLines) { [void]$sb.AppendLine('- ' + $line) }
    }

    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Próxima ação:')
    [void]$sb.AppendLine($Snapshot.Recommendation)
    return $sb.ToString().TrimEnd()
}

function Build-LocalCompareText {
    param($LeftSnapshot, $RightSnapshot)

    $leftScore = $LeftSnapshot.ComparableValue
    $rightScore = $RightSnapshot.ComparableValue
    $winner = 'Empate técnico'
    $winnerReason = 'Os laudos não trouxeram métrica comparável suficientemente clara.'

    if ($null -ne $leftScore -and $null -ne $rightScore) {
        switch ($LeftSnapshot.ComparableMode) {
            'LowerIsBetter' {
                if ([double]$leftScore -lt [double]$rightScore) {
                    $winner = $LeftSnapshot.Label
                    $winnerReason = ("Valor comparável menor: {0} vs {1}" -f $leftScore, $rightScore)
                }
                elseif ([double]$rightScore -lt [double]$leftScore) {
                    $winner = $RightSnapshot.Label
                    $winnerReason = ("Valor comparável menor: {0} vs {1}" -f $rightScore, $leftScore)
                }
            }
        }
    }

    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine('Comparativo:')
    [void]$sb.AppendLine(("Esquerda: {0} | Status: {1} | Hipótese: {2}" -f $LeftSnapshot.Label, $LeftSnapshot.Status, $LeftSnapshot.Hypothesis))
    [void]$sb.AppendLine(("Direita: {0} | Status: {1} | Hipótese: {2}" -f $RightSnapshot.Label, $RightSnapshot.Status, $RightSnapshot.Hypothesis))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine(("Melhor leitura na janela: {0}" -f $winner))
    [void]$sb.AppendLine(("Critério: {0}" -f $winnerReason))
    [void]$sb.AppendLine('')
    [void]$sb.AppendLine('Diferenças principais:')
    if ($null -ne $LeftSnapshot.ComparableValue -or $null -ne $RightSnapshot.ComparableValue) {
        [void]$sb.AppendLine(("- Valor comparável esquerda: {0}" -f [string]$LeftSnapshot.ComparableValue))
        [void]$sb.AppendLine(("- Valor comparável direita: {0}" -f [string]$RightSnapshot.ComparableValue))
    }
    [void]$sb.AppendLine(("- Recomendação esquerda: {0}" -f $LeftSnapshot.Recommendation))
    [void]$sb.AppendLine(("- Recomendação direita: {0}" -f $RightSnapshot.Recommendation))
    return $sb.ToString().TrimEnd()
}

function Build-LocalAskText {
    param($Snapshot, [string]$QuestionText)

    $q = [string]$QuestionText
    if ([string]::IsNullOrWhiteSpace($q)) {
        return 'Pergunta não informada.'
    }

    $qLower = $q.ToLowerInvariant()
    if ($qLower -match 'rede|caminho|tcp|445|ping') {
        if ($Snapshot.Kind -eq 'Single') {
            $dbDown = $Snapshot.Metrics.DbUnavailableSamples
            $netDown = $Snapshot.Metrics.NetDirUnavailableSamples
            if (($dbDown -gt 0) -or ($netDown -gt 0) -or ($Snapshot.Status -eq 'REDE/CAMINHO')) {
                return "Há indício de problema de rede/caminho nesta coleta. A classificação do laudo ficou em $($Snapshot.Status) e houve amostras indisponíveis ou falha objetiva de transporte."
            }
            return "Não há prova forte de quebra de rede/caminho nesta coleta. A leitura principal ficou em $($Snapshot.Status), então a prioridade é validar share/servidor ou pressão local conforme o caso."
        }
        if ($Snapshot.Kind -eq 'Core') {
            $dbDown = $Snapshot.Metrics.DatabaseUnavailableSamples
            $netDown = $Snapshot.Metrics.NetDirUnavailableSamples
            if (($dbDown -gt 0) -or ($netDown -gt 0) -or ($Snapshot.Hypothesis -match 'Compartilhamento')) {
                return "O laudo aponta mais para compartilhamento/caminho do que para software puro, mas a conclusão exata depende do volume de indisponibilidade e da latência observada."
            }
            return "Nesta rodada, não há evidência dominante de quebra de rede. O laudo não registrou sinal forte de indisponibilidade de DB/NetDir."
        }
    }

    if ($qLower -match 'lock|conte.n..o|pdox|netdir') {
        if ($Snapshot.Kind -eq 'Core' -and $Snapshot.Metrics.Contains('PeakLockFileCount')) {
            return ("O pico de lock observado foi {0}. Use isso junto com latência e indisponibilidade para decidir se há contenção real ou apenas baseline nominal." -f [string]$Snapshot.Metrics['PeakLockFileCount'])
        }
        return 'Este formato de laudo não traz uma leitura detalhada de lock equivalente ao core operacional.'
    }

    if ($qLower -match 'cpu|processo|ecgv6|local') {
        if ($Snapshot.Kind -eq 'Core') {
            return ("CPU média/pico: {0}% / {1}%. Processo dominante: {2}" -f [string]$Snapshot.Metrics['AverageCpuPercent'], [string]$Snapshot.Metrics['PeakCpuPercent'], [string]$Snapshot.Metrics['PeakCpuResponsibleSummary'])
        }
        if ($Snapshot.Kind -eq 'Single') {
            return ("CPU média/pico: {0}% / {1}%." -f [string]$Snapshot.Metrics['CpuAvgPercent'], [string]$Snapshot.Metrics['CpuPeakPercent'])
        }
    }

    if ($qLower -match 'pr[oó]ximo|proximo|acao|passo|fazer') {
        return $Snapshot.Recommendation
    }

    return ("Com base no laudo, a leitura atual é {0} com hipótese principal em {1}. Próxima ação sugerida: {2}" -f $Snapshot.Status, $Snapshot.Hypothesis, $Snapshot.Recommendation)
}

function Convert-SnapshotToPromptText {
    param($Snapshot)
    $metricLines = @(Convert-MetricsToLines -Snapshot $Snapshot)
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.AppendLine(("Kind: {0}" -f $Snapshot.Kind))
    [void]$sb.AppendLine(("Label: {0}" -f $Snapshot.Label))
    [void]$sb.AppendLine(("Status: {0}" -f $Snapshot.Status))
    [void]$sb.AppendLine(("Hypothesis: {0}" -f $Snapshot.Hypothesis))
    [void]$sb.AppendLine(("Confidence: {0}" -f $Snapshot.Confidence))
    [void]$sb.AppendLine(("Recommendation: {0}" -f $Snapshot.Recommendation))
    if ($Snapshot.Evidence.Count -gt 0) {
        [void]$sb.AppendLine("Evidence:")
        foreach ($e in $Snapshot.Evidence) { [void]$sb.AppendLine("- " + [string]$e) }
    }
    if ($metricLines.Count -gt 0) {
        [void]$sb.AppendLine("Metrics:")
        foreach ($m in $metricLines) { [void]$sb.AppendLine("- " + [string]$m) }
    }
    return $sb.ToString().TrimEnd()
}

function Invoke-OpenAICompatible {
    param(
        [string]$PromptText,
        $ConfigObject
    )

    if ([string]::IsNullOrWhiteSpace([string]$ConfigObject.Endpoint)) {
        throw 'Endpoint não configurado para provider OpenAICompatible.'
    }
    $apiKeyEnv = [string]$ConfigObject.ApiKeyEnvVar
    if ([string]::IsNullOrWhiteSpace($apiKeyEnv)) { $apiKeyEnv = 'OPENAI_API_KEY' }
    $apiKey = [Environment]::GetEnvironmentVariable($apiKeyEnv, 'Process')
    if ([string]::IsNullOrWhiteSpace($apiKey)) { $apiKey = [Environment]::GetEnvironmentVariable($apiKeyEnv, 'User') }
    if ([string]::IsNullOrWhiteSpace($apiKey)) { $apiKey = [Environment]::GetEnvironmentVariable($apiKeyEnv, 'Machine') }
    if ([string]::IsNullOrWhiteSpace($apiKey)) {
        throw "Variável de ambiente da API não encontrada: $apiKeyEnv"
    }

    $timeoutValue = [int]$ConfigObject.TimeoutSeconds
    if ($TimeoutSeconds -gt 0) { $timeoutValue = $TimeoutSeconds }
    if ($timeoutValue -le 0) { $timeoutValue = 45 }

    $modelName = [string]$ConfigObject.Model
    $temperatureValue = [double]$ConfigObject.Temperature
    $maxTokensValue = [int]$ConfigObject.MaxTokens

    $body = [ordered]@{
        model = $modelName
        temperature = $temperatureValue
        max_tokens = $maxTokensValue
        messages = @(
            @{ role = 'system'; content = 'Você é um analista técnico. Use apenas os fatos fornecidos. Não invente medições.' },
            @{ role = 'user'; content = $PromptText }
        )
    }

    $headers = @{
        'Authorization' = 'Bearer ' + $apiKey
        'Content-Type' = 'application/json'
    }

    $response = Invoke-RestMethod -Method Post -Uri ([string]$ConfigObject.Endpoint) -Headers $headers -Body ($body | ConvertTo-Json -Depth 8) -TimeoutSec $timeoutValue
    if ($null -eq $response -or $null -eq $response.choices -or @($response.choices).Count -eq 0) {
        throw 'Resposta vazia do provider.'
    }
    return [string]$response.choices[0].message.content
}

function Get-ModeOutputName {
    switch ($Mode) {
        'Explain' { return 'ECG_AI_Explain.txt' }
        'Executive' { return 'ECG_AI_Executive.txt' }
        'Technical' { return 'ECG_AI_Technical.txt' }
        'CompareAI' { return 'ECG_AI_Compare.txt' }
        'Ask' { return 'ECG_AI_Ask.txt' }
        default { return 'ECG_AI_Output.txt' }
    }
}

function Get-OutputRootForReport {
    param([string]$PrimaryReportPath, [string]$BaseOutDir)

    if (-not [string]::IsNullOrWhiteSpace($PrimaryReportPath) -and (Test-Path -LiteralPath $PrimaryReportPath)) {
        $reportDir = Split-Path -Parent $PrimaryReportPath
        if ((Split-Path -Leaf $reportDir) -ieq 'AI') {
            return (Split-Path -Parent $reportDir)
        }
        return $reportDir
    }

    Ensure-Directory -Path $BaseOutDir
    $fallback = Join-Path $BaseOutDir ('AI_' + $script:RunStamp)
    Ensure-Directory -Path $fallback
    return $fallback
}

function Save-AiBundle {
    param(
        [string]$PrimaryReportPath,
        [string]$BaseOutDir,
        [string]$RenderedText,
        [string]$PromptText,
        [string]$RawResponseText,
        [object]$MetaObject
    )

    $outputRoot = Get-OutputRootForReport -PrimaryReportPath $PrimaryReportPath -BaseOutDir $BaseOutDir
    $aiDir = Join-Path $outputRoot 'AI'
    Ensure-Directory -Path $aiDir
    $script:CurrentOutDir = $outputRoot
    $script:CurrentAiDir = $aiDir

    $mainFile = Join-Path $aiDir (Get-ModeOutputName)
    $metaFile = Join-Path $aiDir 'ECG_AI_Metadata.json'
    Write-Utf8File -Path $mainFile -Text $RenderedText
    Write-Utf8File -Path $metaFile -Text ($MetaObject | ConvertTo-Json -Depth 8)

    if ($SavePromptCopy) {
        Write-Utf8File -Path (Join-Path $aiDir 'ECG_AI_Prompt.txt') -Text $PromptText
    }
    if ($SaveRawResponse) {
        Write-Utf8File -Path (Join-Path $aiDir 'ECG_AI_RawResponse.txt') -Text $RawResponseText
    }

    return [PSCustomObject]@{
        MainFile = $mainFile
        MetaFile = $metaFile
        AiDir = $aiDir
    }
}

function Save-AiError {
    param([string]$BaseOutDir, [string]$ErrorText)
    $root = Get-OutputRootForReport -PrimaryReportPath $ReportPath -BaseOutDir $BaseOutDir
    $aiDir = Join-Path $root 'AI'
    Ensure-Directory -Path $aiDir
    $script:CurrentAiDir = $aiDir
    $path = Join-Path $aiDir 'ECG_AI_Error.log'
    $payload = @(
        "ToolVersion=$($script:ToolVersion)",
        "Mode=$Mode",
        "Error=$ErrorText",
        ($script:LogLines -join [Environment]::NewLine)
    ) -join [Environment]::NewLine
    Write-Utf8File -Path $path -Text $payload
    return $path
}

try {
    $resolvedProfile = Resolve-ProfilePath
    $effectiveOutDir = Resolve-EffectiveOutDir -ResolvedProfilePath $resolvedProfile
    $resolvedConfigPath = Resolve-ConfigPath
    $resolvedPromptPath = Resolve-PromptCatalogPath

    $config = Read-JsonConfigOrDefault -Path $resolvedConfigPath
    $prompts = Read-PromptCatalogOrDefault -Path $resolvedPromptPath

    if (-not [string]::IsNullOrWhiteSpace($Provider)) { $config.Provider = $Provider }
    if ($TimeoutSeconds -gt 0) { $config.TimeoutSeconds = $TimeoutSeconds }
    if ($NoRedaction) { $config.RedactSensitiveData = $false }
    if ($SavePromptCopy) { $config.SavePromptCopy = $true }
    if ($SaveRawResponse) { $config.SaveRawResponse = $true }

    Log "$($script:ToolName) $($script:ToolVersion)"
    Log "Mode: $Mode"
    Log "Profile: $resolvedProfile"
    Log "OutDir: $effectiveOutDir"
    Log "Config: $resolvedConfigPath"
    Log "Prompts: $resolvedPromptPath"
    Log "Provider: $($config.Provider)"

    $primaryReportPath = ''
    $leftSnapshot = $null
    $rightSnapshot = $null
    $snapshot = $null

    switch ($Mode) {
        'CompareAI' {
            if ([string]::IsNullOrWhiteSpace($LeftReportPath) -or [string]::IsNullOrWhiteSpace($RightReportPath)) {
                throw 'CompareAI requer -LeftReportPath e -RightReportPath.'
            }
            $LeftReportPath = Resolve-ReportInputPath -InputPath $LeftReportPath
            $RightReportPath = Resolve-ReportInputPath -InputPath $RightReportPath
            $leftData = Read-JsonFile -Path $LeftReportPath
            $rightData = Read-JsonFile -Path $RightReportPath
            if ($config.RedactSensitiveData) {
                $leftData = ConvertTo-RedactedObject -InputObject $leftData
                $rightData = ConvertTo-RedactedObject -InputObject $rightData
            }
            $leftSnapshot = Build-NormalizedSnapshot -Data $leftData -SourcePath $LeftReportPath
            $rightSnapshot = Build-NormalizedSnapshot -Data $rightData -SourcePath $RightReportPath
            $primaryReportPath = $LeftReportPath
        }
        default {
            if ([string]::IsNullOrWhiteSpace($ReportPath)) {
                $ReportPath = Find-LatestReportPath -BaseOutDir $effectiveOutDir
            } else {
                $ReportPath = Resolve-ReportInputPath -InputPath $ReportPath
            }
            if ($Mode -eq 'Ask' -and $Question -match '^\s*\d+\s*$') {
                throw 'Pergunta invalida: foi recebido apenas um numero. Informe uma pergunta real sobre o laudo.'
            }
            $data = Read-JsonFile -Path $ReportPath
            if ($config.RedactSensitiveData) {
                $data = ConvertTo-RedactedObject -InputObject $data
            }
            $snapshot = Build-NormalizedSnapshot -Data $data -SourcePath $ReportPath
            $primaryReportPath = $ReportPath
        }
    }

    $rendered = ''
    $promptText = ''
    $rawResponse = ''
    $usedProvider = [string]$config.Provider
    $fallbackUsed = $false

    switch ($Mode) {
        'Explain' {
            $promptText = ($prompts.Explain + [Environment]::NewLine + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $snapshot))
            if ($usedProvider -eq 'LocalRules') {
                $rawResponse = Build-LocalExplainText -Snapshot $snapshot
                $rendered = $rawResponse
            } else {
                try {
                    $rawResponse = Invoke-OpenAICompatible -PromptText $promptText -ConfigObject $config
                    $rendered = $rawResponse
                } catch {
                    if (-not $config.FallbackWithoutAI) { throw }
                    $fallbackUsed = $true
                    $rawResponse = Build-LocalExplainText -Snapshot $snapshot
                    $rendered = $rawResponse
                }
            }
        }
        'Executive' {
            $promptText = ($prompts.Executive + [Environment]::NewLine + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $snapshot))
            if ($usedProvider -eq 'LocalRules') {
                $rawResponse = Build-LocalExecutiveText -Snapshot $snapshot
                $rendered = $rawResponse
            } else {
                try {
                    $rawResponse = Invoke-OpenAICompatible -PromptText $promptText -ConfigObject $config
                    $rendered = $rawResponse
                } catch {
                    if (-not $config.FallbackWithoutAI) { throw }
                    $fallbackUsed = $true
                    $rawResponse = Build-LocalExecutiveText -Snapshot $snapshot
                    $rendered = $rawResponse
                }
            }
        }
        'Technical' {
            $promptKey = 'Technical_' + $Audience
            $promptSeed = [string]$prompts.$promptKey
            if ([string]::IsNullOrWhiteSpace($promptSeed)) { $promptSeed = [string]$prompts.Technical_INFRA }
            $promptText = ($promptSeed + [Environment]::NewLine + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $snapshot))
            if ($usedProvider -eq 'LocalRules') {
                $rawResponse = Build-LocalTechnicalText -Snapshot $snapshot -Audience $Audience
                $rendered = $rawResponse
            } else {
                try {
                    $rawResponse = Invoke-OpenAICompatible -PromptText $promptText -ConfigObject $config
                    $rendered = $rawResponse
                } catch {
                    if (-not $config.FallbackWithoutAI) { throw }
                    $fallbackUsed = $true
                    $rawResponse = Build-LocalTechnicalText -Snapshot $snapshot -Audience $Audience
                    $rendered = $rawResponse
                }
            }
        }
        'CompareAI' {
            $promptText = ($prompts.CompareAI + [Environment]::NewLine + [Environment]::NewLine +
                "ESQUERDA" + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $leftSnapshot) + [Environment]::NewLine + [Environment]::NewLine +
                "DIREITA" + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $rightSnapshot))
            if ($usedProvider -eq 'LocalRules') {
                $rawResponse = Build-LocalCompareText -LeftSnapshot $leftSnapshot -RightSnapshot $rightSnapshot
                $rendered = $rawResponse
            } else {
                try {
                    $rawResponse = Invoke-OpenAICompatible -PromptText $promptText -ConfigObject $config
                    $rendered = $rawResponse
                } catch {
                    if (-not $config.FallbackWithoutAI) { throw }
                    $fallbackUsed = $true
                    $rawResponse = Build-LocalCompareText -LeftSnapshot $leftSnapshot -RightSnapshot $rightSnapshot
                    $rendered = $rawResponse
                }
            }
        }
        'Ask' {
            $promptText = ($prompts.Ask + [Environment]::NewLine + [Environment]::NewLine +
                "Pergunta: " + $Question + [Environment]::NewLine + [Environment]::NewLine + (Convert-SnapshotToPromptText -Snapshot $snapshot))
            if ($usedProvider -eq 'LocalRules') {
                $rawResponse = Build-LocalAskText -Snapshot $snapshot -QuestionText $Question
                $rendered = $rawResponse
            } else {
                try {
                    $rawResponse = Invoke-OpenAICompatible -PromptText $promptText -ConfigObject $config
                    $rendered = $rawResponse
                } catch {
                    if (-not $config.FallbackWithoutAI) { throw }
                    $fallbackUsed = $true
                    $rawResponse = Build-LocalAskText -Snapshot $snapshot -QuestionText $Question
                    $rendered = $rawResponse
                }
            }
        }
    }

    $meta = [PSCustomObject]@{
        ToolName = $script:ToolName
        ToolVersion = $script:ToolVersion
        GeneratedAt = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        Mode = $Mode
        ProviderRequested = $usedProvider
        FallbackUsed = $fallbackUsed
        ReportPath = $ReportPath
        LeftReportPath = $LeftReportPath
        RightReportPath = $RightReportPath
        Audience = $Audience
        Question = $Question
        OutDir = $effectiveOutDir
        ConfigPath = $resolvedConfigPath
        PromptCatalogPath = $resolvedPromptPath
    }

    $saved = Save-AiBundle -PrimaryReportPath $primaryReportPath -BaseOutDir $effectiveOutDir -RenderedText $rendered -PromptText $promptText -RawResponseText $rawResponse -MetaObject $meta

    Log "Saída principal salva em $($saved.MainFile)" 'STEP'
    Log "Metadados salvos em $($saved.MetaFile)" 'STEP'

    if ($OpenOutput) {
        try { Start-Process $saved.MainFile | Out-Null } catch {}
    }

    exit 0
}
catch {
    $msg = $_.Exception.Message
    Log "ERRO FATAL: $msg" 'ERROR'
    try {
        $errorPath = Save-AiError -BaseOutDir (Resolve-EffectiveOutDir -ResolvedProfilePath (Resolve-ProfilePath)) -ErrorText $msg
        Log "Log de erro salvo em $errorPath" 'STEP'
    } catch {}
    exit 1
}
