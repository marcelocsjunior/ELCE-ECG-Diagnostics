# 02_COMPARE_EXECUTANTE_VIEWER

## Pré-requisitos

- um laudo da viewer com `StationRole=VIEWER`
- um laudo da executante com `StationRole=EXECUTANTE`
- ambos apontando para o mesmo DB e NetDir esperados

## Objetivo do compare

Comprovar convergência estrutural entre estações com papéis distintos sem contaminar a evidência por heurística.

## Deve convergir

- caminho efetivo do DB
- NetDir esperado e efetivo
- convergência BDE/NETDIR
- integridade do `IDAPI32.CFG`
- capacidade de write probe no NetDir

## Pode divergir de forma esperada

- papel da estação
- alias da estação
- executável local e contexto operacional específico
- detalhes de usuário/perfil/logon

## Condição de aptidão

Resultado final desejado:
- `CONVERGENTE`
- `APTO_PARA_PILOTO`
