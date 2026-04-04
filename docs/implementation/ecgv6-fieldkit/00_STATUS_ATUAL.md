# 00_STATUS_ATUAL

## Resumo executivo

O ECGv6 FieldKit foi validado operacionalmente em uma estação viewer.
Ainda não houve validação real da estação executante e ainda não houve compare real entre executante e viewer.

## Comprovado na viewer

- DB acessível
- NetDir acessível
- WriteProbe funcional
- locks estáveis com pico 2
- convergência BDE/NETDIR em HKLM, WOW6432Node e HKCU
- `IDAPI32.CFG` corrigido para `\\192.168.1.57\Database\NetDir`
- backup `.bak` gerado
- HTML sem mojibake
- compare funcional entre duas rodadas da mesma viewer

## Pendências críticas

1. validar executante real
2. ajustar operação com `StationRole` explícito
3. executar compare real executante x viewer
4. decidir promoção para `main` após evidência

## Decisão de engenharia

- manter o core principal intacto em `src/`
- manter o FieldKit isolado em `fixpacks/ecgv6-fieldkit/`
- promover primeiro como preview/pilot track
