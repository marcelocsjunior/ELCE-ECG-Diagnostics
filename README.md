# ELCE ECG Diagnostics

## Status oficial

**Baseline oficial em producao:** `ECG Diagnostics Suite v6.3.2` / `v6.3.2-hardening`.

**Baseline historica homologada anterior:** `ECG Diagnostics Core v5.2.4` / `5.2.4-unified-stable`.

Este repositorio reflete a **baseline oficial atual em producao** na linha `package/v6.3.2/` e preserva a linha `v5.2.4` como historico homologado anterior.

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

- `package/v6.3.2/` = baseline oficial vigente
- `package/v5.2.4/` = baseline historica homologada anterior
- a linha oficial promovida usa a suite v6.3.2 com backend WS2016 como alvo oficial, mantendo FS01 e XP para A/B ou contingencia conforme a release

## O que vale como verdade operacional

- Core oficial: `ECG_Diagnostics_Core_v6_3_2.ps1`
- Launcher oficial: `ECG_Diagnostics_Hub_v6_3_2.bat`
- Compare launcher oficial: `ECG_CompareBackend_Launcher_v6_3_2.bat`
- Builder oficial: `ECG_ProfileBuilder_v6_3_2.ps1`
- Perfil oficial: `ECG_FieldKit_Unified_v6_3_2.ini`
- Release note do pacote: `README_v6_3_2.txt`
- Hashes de integridade: `SHA256SUMS_v6_3_2.txt`

## Modos suportados na baseline oficial

- `Fix`
- `Auto`
- `Detect`
- `Compare`
- `CompareJson`
- `Rollback`
- `Monitor`
- `CollectStatic`
- `Single`
- `CompareBackend`

## Saidas esperadas da baseline oficial

```text
C:\ECG\FieldKit\out\<RunId>\ECG_Report.html
C:\ECG\FieldKit\out\<RunId>\ECG_Report.json
C:\ECG\FieldKit\out\<RunId>\ECG_Fatal_Error.log
```

## Perfil INI oficial

O perfil oficial atual esta em `package/v6.3.2/ECG_FieldKit_Unified_v6_3_2.ini`.

Presets oficiais da linha promovida:
- FS01 legado por hostname
- XP legado por IP
- WS2016 novo oficial
- Custom

## Pre-requisitos de execucao

- Windows com **Windows PowerShell 5.1** disponivel
- acesso ao compartilhamento UNC do banco e do `NetDir`
- acesso ao executavel do ECG em `C:\HW\ECG\ECGV6.exe` ou aderencia ao caminho configurado no perfil
- permissao de escrita em `C:\ECG\FieldKit\out`
- para `Fix` e `Rollback`, execucao **elevada** (Administrador)
- para cenarios com BDE, acesso ao registro e, quando aplicavel, ao `IDAPI32.CFG`

## Regra pratica de uso

- **Primeiro `Auto`** para entender o incidente sem interferencia
- **Depois `Fix`** somente quando a causa provavel estiver clara e houver autorizacao
- **`Single`** para diagnostico de um target especifico
- **`CompareBackend`** para comparacao estruturada entre os backends habilitados

## Diretriz de governanca

- **v6.3.2 e a baseline oficial em producao**.
- **v5.2.4 permanece como baseline historica homologada anterior**.
- A linha antiga em `src/` permanece somente como referencia historica enquanto a migracao documental nao for concluida.
- Toda evolucao futura deve partir de `package/v6.3.2/` ou de sua proxima baseline promovida.

## Nota de validacao

A validacao final de qualquer alteracao futura deve ocorrer em Windows real, preferencialmente em estacao representativa do ambiente.

## Nota sobre legado

O conteudo anterior do repositorio continua existindo para historico tecnico, mas **nao e a baseline operacional atual**.
