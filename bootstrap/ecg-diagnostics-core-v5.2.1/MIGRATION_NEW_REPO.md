# Plano de migração para o novo repositório

## Estratégia recomendada

1. Criar um repositório novo e independente para a ferramenta.
2. Subir somente a árvore do produto standalone.
3. Tratar a primeira publicação como `v5.2.1`.
4. Manter o repositório atual apenas como baseline histórica e trilha anterior.

## Estrutura sugerida

```text
/core/ECG_Diagnostics_Core.ps1
/hub/ECG_Diagnostics_Hub.bat
/config/ECG_FieldKit.ini.example
/docs/README_RELEASE.md
/docs/ROLLBACK.md
/releases/SHA256SUMS.txt
```

## Branching sugerido

- `main` -> linha estável
- `develop` -> opcional, apenas se houver evolução paralela
- hotfixes pequenos direto por branch curta

## Política de release

- tag inicial: `v5.2.1`
- release notes derivadas do `README_RELEASE.txt`
- checksums publicados junto da release
- recomendação inicial: manter o novo repositório privado até a primeira validação controlada em Windows real

## Regra de corte

Não arrastar para o novo repositório:
- histórico de PRs de fieldkit do repo legado
- documentação de baseline read-only antiga que não pertença ao produto atual
- artefatos congelados do produto anterior
