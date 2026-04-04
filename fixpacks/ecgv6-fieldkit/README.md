# ECGv6 FieldKit

Trilha segregada de diagnóstico e remediação controlada para cenários legados de ECGv6/BDE/NETDIR.

## Governança

- **Não substitui** o core principal em `src/`
- **Não altera** a baseline read-only por padrão
- Deve ser tratado como **fix pack especializado** sob `fixpacks/`
- Evidência oficial deve usar `StationRole` explícito

## Localização canônica

- `fixpacks/ecgv6-fieldkit/ECGv6_FieldKit.ps1`
- `fixpacks/ecgv6-fieldkit/ECGv6_FieldKit.bat`
- `fixpacks/ecgv6-fieldkit/ECGv6_FieldKit.ini.example`
- `fixpacks/ecgv6-fieldkit/profiles/*.ini`

## Modos suportados

- `Prepare`
- `Audit`
- `Auto`
- `Fix`
- `Compare`
- `Rollback`

## Estado atual da validação

Validado operacionalmente em **uma estação viewer**.

Comprovado na viewer:
- acesso ao DB em `\\192.168.1.57\Database`
- acesso ao NetDir em `\\192.168.1.57\Database\NetDir`
- `WriteProbe` funcional no NetDir
- locks estáveis com pico 2
- convergência BDE/NETDIR em `HKLM`, `WOW6432Node` e `HKCU`
- `IDAPI32.CFG` corrigido de `C:\HW\Database\NetDir` para `\\192.168.1.57\Database\NetDir`
- backup `.bak` do `IDAPI32.CFG`
- HTML íntegro, sem mojibake
- compare funcional entre duas rodadas da mesma viewer

Ainda pendente:
- validação real da estação executante
- compare real executante x viewer

## Perfis recomendados

- `profiles/viewer.ini`
- `profiles/executante.ini`
- `profiles/host_xp.ini`

**Regra operacional:** evitar `StationRole=AUTO` em piloto oficial, evidência e compare.

## Fluxo recomendado

1. Validar executante com `StationRole=EXECUTANTE`
2. Confirmar estabilidade de DB/NetDir/BDE/IDAPI32.CFG
3. Rodar compare real entre executante e viewer
4. Consolidar evidências
5. Só então promover PR para `main`

## Observação

O repositório principal continua com o produto ELCE ECG Diagnostics como baseline read-only.
O FieldKit entra como trilha paralela de remediação controlada.
