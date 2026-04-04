## Bloco sugerido para inserir no README.md

### ECGv6 FieldKit

O ECGv6 FieldKit é uma trilha separada de diagnóstico e remediação controlada para cenários legados de ECGv6/BDE/NETDIR.

Ele não substitui o core principal em `src/` e não deve ser tratado como parte do fluxo read-only de laudo operacional. Sua função é atuar como fix pack especializado, com execução controlada, geração de evidências e capacidade de compare e rollback.

#### Localização canônica

- `fixpacks/ecgv6-fieldkit/` -> motor, launcher e perfis do FieldKit
- `docs/runbooks/runbookECGv6-FieldKit.txt` -> runbook operacional
- `docs/implementation/ecgv6-fieldkit/` -> documentação de validação, compare e critérios de aceitação

#### Escopo do FieldKit

- `Prepare`
- `Audit`
- `Auto`
- `Fix`
- `Compare`
- `Rollback`

#### Status atual

- validado operacionalmente em uma estação viewer
- validação da estação executante real ainda pendente
- compare real entre executante e viewer ainda pendente
- uso recomendado com `StationRole` explícito por perfil (`VIEWER`, `EXECUTANTE`, `HOST_XP`)

#### Regra de governança

O FieldKit permanece em trilha segregada dentro de `fixpacks/`.
A baseline read-only do produto continua em `src/`.
Nenhuma mudança do FieldKit altera por padrão o comportamento do core principal.
