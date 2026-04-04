# 01_ESTRATEGIA_VALIDACAO

## Objetivo
Fechar a lacuna entre validação parcial da viewer e readiness de promoção do FieldKit.

## Ordem correta

1. validar estação executante real
2. consolidar `StationRole` explícito
3. comparar executante x viewer
4. consolidar documentação e release note
5. abrir PR draft / promover branch

## Configuração obrigatória
- `ExpectedDbPath=\\192.168.1.57\Database`
- `ExpectedNetDir=\\192.168.1.57\Database\NetDir`
- `ExpectedExePath=C:\HW\ECG\ECGV6.exe`

## Papel por estação
- viewer -> `StationRole=VIEWER`
- executante -> `StationRole=EXECUTANTE`
- host XP -> `StationRole=HOST_XP`

## Regra
`AUTO` só deve existir por compatibilidade retroativa.
Para evidência formal, laudo homologatório e compare, o papel deve ser explícito.

## Critério de aprovação da executante
- DB e NetDir acessíveis
- WriteProbe OK
- locks estáveis
- BDE convergente nas hives relevantes
- `IDAPI32.CFG` convergente
- laudo HTML íntegro
- operação funcional do ECGV6 sem regressão
