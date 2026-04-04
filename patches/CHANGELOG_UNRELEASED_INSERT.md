## Bloco sugerido para inserir em CHANGELOG.md / [Unreleased]

### Added
- Adicionada a trilha segregada `fixpacks/ecgv6-fieldkit/` para o ECGv6 FieldKit.
- Adicionados perfis explícitos `VIEWER`, `EXECUTANTE` e `HOST_XP` para reduzir ambiguidade de evidência e compare.
- Adicionado runbook dedicado `docs/runbooks/runbookECGv6-FieldKit.txt`.
- Adicionada documentação de implementação em `docs/implementation/ecgv6-fieldkit/`.
- Adicionada release note de preview `releases/notes/v1.2.0-fieldkit-preview.md`.

### Changed
- Formalizada a governança do FieldKit como trilha paralela de remediação, sem alteração do core read-only em `src/`.

### Notes
- Validação atual cobre apenas uma estação viewer.
- Validação da executante real e compare executante x viewer permanecem pendentes antes da promoção final.
