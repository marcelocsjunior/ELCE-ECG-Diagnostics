# 06_LIMITACOES_E_RISCOS

## Limitações atuais

- problemas residuais de encoding / acentuação;
- ajustes pendentes no gráfico HTML;
- classificação de tipo de máquina ainda incompleta;
- sem análise Defender/minifilter;
- sem benchmark ativo;
- sem comparação automática com referência.

## Riscos atuais

1. promoção incorreta de build intermediária;
2. mistura de BAT e PS1 de linhas diferentes;
3. perda da golden copy;
4. correção pontual que reabre regressão no fechamento da rodada.

## Postura recomendada

- patch mínimo;
- rollback simples;
- validação fechada;
- sem refactor estrutural prematuro.
