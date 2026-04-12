# ELCE ECG Diagnostics

## Status oficial

**Baseline oficial em producao:** `ECG Diagnostics Core v5.2.4` / `5.2.4-unified-stable`.

**Baseline candidata publicada no repositorio:** `v6.3.2-hardening` em `package/v6.3.2/`.

Este repositorio reflete o **pacote homologado em producao** e tambem a **proxima baseline candidata** em trilha de homologacao, mantendo a linha anterior apenas como historico/legado.

## Pacotes versionados

```text
package/v5.2.4/
  ECG_Diagnostics_Core.ps1
  ECG_Diagnostics_Hub.bat
  ECG_FieldKit.ini
  README_RELEASE.txt
  SHA256SUMS.txt

package/v6.3.2/
  ECG_Diagnostics_Core_v6_3_2.ps1
  ECG_Diagnostics_Hub_v6_3_2.bat
  ECG_CompareBackend_Launcher_v6_3_2.bat
  ECG_ProfileBuilder_v6_3_2.ps1
  ECG_FieldKit_Unified_v6_3_2.ini
  ECG_AI_Prototype_Hub_v0_2_1.bat
  ECG_Diagnostics_AI_Prototype_v0_2_1.ps1
  ECG_AI.config.json
  ECG_AI_Prompts.json
  ECG_v6_3_2_Debug_Checklist.txt
  HOTFIX_6_3_2_NOTES.txt
  README_v6_3_2.txt
  SHA256SUMS_v6_3_2.txt
```

## Leitura rapida para quem abre o repo hoje

- `package/v5.2.4/` = baseline oficial vigente
- `package/v6.3.2/` = baseline candidata / promotavel
- a promocao final da `v6.3.2` depende de homologacao funcional em Windows PowerShell 5.1 real

## O que vale como verdade operacional

- Core oficial: `ECG_Diagnostics_Core.ps1`
- Launcher oficial: `ECG_Diagnostics_Hub.bat`
- Perfil oficial: `ECG_FieldKit.ini`
- Release note do pacote: `README_RELEASE.txt`
- Hashes de integridade: `SHA256SUMS.txt`

## Modos suportados na baseline oficial

- `Fix`
- `Auto`
- `Compare`
- `Rollback`
- `Monitor`
- `CollectStatic`

## Saidas esperadas da baseline oficial

```text
C:\ECG\FieldKit\out\<RunId>\ECG_Report.html
C:\ECG\FieldKit\out\<RunId>\ECG_Report.json
C:\ECG\FieldKit\out\<RunId>\ECG_Fatal_Error.log
```

## Perfil INI oficial

O perfil oficial atual esta em `package/v5.2.4/ECG_FieldKit.ini` e define o contrato basico da estacao.

```ini
ExpectedDbPath=\\192.168.1.57\Database
ExpectedNetDir=\\192.168.1.57\Database\NetDir
ExpectedExePath=C:\HW\ECG\ECGV6.exe
SetMachineHwPath=true
OutDir=C:\ECG\FieldKit\out
StationRole=AUTO
StationAlias=VIEWER
MonitorMinutes=3
SampleIntervalSeconds=15
CpuProcessCaptureThreshold=80
TopProcessCaptureCount=3
EnableLatencyMetrics=true
EnableEcgProcessMetrics=true
EnableDiskMetrics=false
EnableNetworkMetrics=false
```

Leitura rapida do perfil oficial:
- caminho esperado do banco em UNC `\\192.168.1.57\Database`
- `NetDir` esperado em `\\192.168.1.57\Database\NetDir`
- executavel esperado em `C:\HW\ECG\ECGV6.exe`
- output padrao em `C:\ECG\FieldKit\out`
- latencia UNC e CPU do processo ECGV6 habilitadas
- fila de disco e bytes de rede opcionais desabilitados por padrao

## Pre-requisitos de execucao

- Windows com **Windows PowerShell 5.1** disponivel
- acesso ao compartilhamento UNC do banco e do `NetDir`
- acesso ao executavel do ECG em `C:\HW\ECG\ECGV6.exe` ou aderencia ao caminho configurado no perfil
- permissao de escrita em `C:\ECG\FieldKit\out`
- para `Fix` e `Rollback`, execucao **elevada** (Administrador)
- para cenarios com BDE, acesso ao registro e, quando aplicavel, ao `IDAPI32.CFG`

## Quando usar `Auto` vs `Fix`

### `Auto`
Use `Auto` quando o objetivo for **somente diagnostico**, sem tocar em registro, variaveis ou `IDAPI32.CFG`.

Cenario tipico:
- validar sintoma reportado
- medir latencia UNC
- capturar CPU dominante
- gerar laudo HTML/JSON sem mudanca no ambiente

### `Fix`
Use `Fix` quando ja houver evidencia suficiente de desalinhamento de `NETDIR`, variavel `HW_CAMINHO_DB` ou `IDAPI32.CFG`, e a acao corretiva estiver autorizada.

Cenario tipico:
- corrigir `NETDIR` em `HKLM/HKCU`
- alinhar `IDAPI32.CFG`
- gerar backup de rollback `.reg`
- aplicar ajuste controlado e emitir laudo da rodada

### Regra pratica
- **Primeiro `Auto`**, para entender o incidente sem interferencia.
- **Depois `Fix`**, somente quando a causa provavel estiver clara e houver janela/autorizacao para remediacao.

## Diretriz de governanca

- **v5.2.4 e a baseline oficial em producao**.
- A linha antiga em `src/` permanece somente como referencia historica enquanto a migracao documental nao for concluida.
- Nao misturar launcher/core da linha antiga com o pacote oficial v5.2.4.
- Toda evolucao futura deve partir de `package/v5.2.4/` ou de sua proxima baseline promovida.

## Validacao minima antes de promover alteracoes futuras

1. Validar `Auto`
2. Validar `Monitor`
3. Validar `Fix` elevado
4. Validar `Compare`
5. Validar `Rollback`
6. Validar geracao de `ECG_Report.html`
7. Validar geracao de `ECG_Report.json`
8. Validar ausencia de `ECG_Fatal_Error.log` em rodada bem sucedida

## Nota de validacao

Embora o pacote esteja homologado em producao, a **validacao final de qualquer alteracao futura deve ocorrer em Windows real**, preferencialmente em estacao representativa do ambiente.

Este repositorio pode ser analisado e versionado fora do Windows, mas isso **nao substitui** teste funcional real com:
- acesso ao UNC
- acesso ao ECGV6
- leitura/escrita de registro
- leitura/ajuste de `IDAPI32.CFG`
- verificacao do HTML e do JSON gerados

## Nota sobre legado

O conteudo anterior do repositorio continua existindo para historico tecnico, mas **nao e a baseline operacional atual**.
