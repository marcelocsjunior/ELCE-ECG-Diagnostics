# ELCE ECG Diagnostics

## Status oficial

**Baseline oficial em producao:** `ECG Diagnostics Core v5.2.4` / `5.2.4-unified-stable`.

Este repositorio passa a refletir o **pacote homologado em producao**, mantendo a linha anterior apenas como historico/legado.

## Pacote oficial versionado

```text
package/v5.2.4/
  ECG_Diagnostics_Core.ps1
  ECG_Diagnostics_Hub.bat
  ECG_FieldKit.ini
  README_RELEASE.txt
  SHA256SUMS.txt
```

## O que vale como verdade operacional

- Core oficial: `ECG_Diagnostics_Core.ps1`
- Launcher oficial: `ECG_Diagnostics_Hub.bat`
- Perfil oficial: `ECG_FieldKit.ini`
- Release note do pacote: `README_RELEASE.txt`

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

## Nota sobre legado

O conteudo anterior do repositorio continua existindo para historico tecnico, mas **nao e a baseline operacional atual**.

