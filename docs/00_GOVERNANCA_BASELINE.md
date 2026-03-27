# 00_GOVERNANCA_BASELINE

## Princípio central

A baseline operacional restaurada é a **golden copy** do produto.

## Regras obrigatórias

- Não misturar BAT de uma linha com PS1 de outra.
- Não promover builds intermediárias só porque possuem um ajuste desejável.
- Qualquer patch futuro deve partir da baseline válida.
- Qualquer patch futuro deve ser mínimo, reversível e validado isoladamente.
- Staging e produção só podem divergir de forma consciente, documentada e rastreável.

## Fonte operacional atual

Par operacional:
- `src/ELCE_ECG_Diagnostics.ps1`
- `src/ELCE_ECG_Diagnostics_Menu.bat`

## Risco principal nesta fase

O risco atual não é mais “o HTML não nasce”.  
O risco principal é **promover linha errada** e perder a cópia funcional.

## Regra de promoção

Uma mudança só pode subir de baseline se:
1. preservar compatibilidade BAT + PS1;
2. não quebrar fechamento da rodada;
3. mantiver o HTML principal;
4. mantiver os JSONs obrigatórios;
5. passar no checklist de validação oficial.
