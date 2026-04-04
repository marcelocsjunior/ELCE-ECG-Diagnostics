# MOTOR_SYNC_STATUS

## Estado atual

Este branch já contém a trilha estrutural do ECGv6 FieldKit no repositório:

- launcher BAT
- perfil `.ini` de exemplo
- perfis explícitos (`VIEWER`, `EXECUTANTE`, `HOST_XP`)
- runbook
- documentação de status, validação, compare e aceite
- release note de preview
- snippets de patch para `README.md` e `CHANGELOG.md`

## Observação importante

O motor PowerShell final `ECGv6_FieldKit.ps1` permanece preservado no pacote local preparado para sincronização completa.

Pacote local preparado nesta sessão:
- `ELCE-ECG-FieldKit-PR-Package.zip`

## Motivo desta sinalização

O wrapper do conector disponível nesta sessão permitiu criação de branch, novos arquivos e PR, mas não expôs um fluxo prático para sincronizar com segurança o motor local final de grande porte e, ao mesmo tempo, atualizar arquivos já existentes como `README.md` e `CHANGELOG.md` diretamente no branch.

## Decisão operacional

Este PR deve permanecer em **draft** até:
1. sincronização do motor `ECGv6_FieldKit.ps1`
2. validação da executante real
3. compare real executante x viewer
4. atualização final de `README.md` e `CHANGELOG.md`
