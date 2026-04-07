# Release manifest — ECG Diagnostics Core v5.2.1

## Pacote lógico

- ECG_Diagnostics_Core.ps1
- ECG_Diagnostics_Hub.bat
- ECG_FieldKit.ini
- README_RELEASE.txt
- SHA256SUMS.txt

## Capacidades funcionais consolidadas

- Fix
- Auto
- Compare
- Rollback
- Monitor
- CollectStatic

## Implantação alvo

```text
C:\ECG\FieldKit
```

## Saídas esperadas

### Auto / Fix / Monitor
- `C:\ECG\FieldKit\out\<RunId>\ECG_Report.html`
- `C:\ECG\FieldKit\out\<RunId>\ECG_Report.json`
- `C:\ECG\FieldKit\out\<RunId>\ECG_Fatal_Error.log` (somente em falha)

### CollectStatic
- `C:\ECG\FieldKit\out\static\ECG_State_<HOST>_<TIMESTAMP>.json`
- `C:\ECG\FieldKit\out\static\ECG_Fatal_Error.log` (somente em falha)

## Nota honesta de validação

A base foi revisada e empacotada de forma consistente, mas ainda depende de validação controlada em Windows real antes de rollout amplo.
