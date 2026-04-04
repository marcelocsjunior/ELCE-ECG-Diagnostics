# 03_MATRIZ_DE_ACEITACAO

| Item | Viewer | Executante | Compare |
|---|---|---|---|
| DB acessível | obrigatório | obrigatório | deve convergir |
| NetDir acessível | obrigatório | obrigatório | deve convergir |
| WriteProbe | obrigatório | obrigatório | deve convergir |
| Locks estáveis | obrigatório | obrigatório | deve convergir |
| BDE HKLM | obrigatório | obrigatório | deve convergir |
| BDE WOW6432Node | obrigatório | obrigatório | deve convergir |
| BDE HKCU | obrigatório | obrigatório | deve convergir |
| IDAPI32.CFG | obrigatório | obrigatório | deve convergir |
| Role explícito | VIEWER | EXECUTANTE | obrigatório |
| HTML íntegro | obrigatório | obrigatório | obrigatório |

## Gate de promoção

O PR do FieldKit só deve sair de draft quando:
1. a executante real estiver validada
2. o compare executante x viewer estiver concluído
3. a documentação estiver alinhada
