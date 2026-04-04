# Changelog

Todas as mudanças relevantes deste produto devem ser registradas aqui.

O formato segue uma linha inspirada em *Keep a Changelog* e versionamento semântico.

## [Unreleased]

### Added
- Adicionada a trilha segregada `fixpacks/ecgv6-fieldkit/` para o ECGv6 FieldKit.
- Adicionados perfis explícitos `VIEWER`, `EXECUTANTE` e `HOST_XP` para reduzir ambiguidade de evidência e compare.
- Adicionado o runbook dedicado `docs/runbooks/runbookECGv6-FieldKit.txt`.
- Adicionada documentação de implementação em `docs/implementation/ecgv6-fieldkit/`.
- Adicionada a release note de preview `releases/notes/v1.2.0-fieldkit-preview.md`.

### Changed
- Formalizada a governança do FieldKit como trilha paralela de remediação, sem alteração do core read-only em `src/`.
- Documentado no `README.md` que o FieldKit permanece fora do fluxo principal de laudo.

### Notes
- A validação atual cobre apenas uma estação viewer.
- A validação da executante real e o compare executante x viewer permanecem pendentes antes da promoção final.

## [1.1.0] - 2026-03-28

### Added
- Suporte operacional multiunidade para UN1, UN2 e UN3 por perfil.
- Inclusão de `ECG_UnitProfiles.json` para centralizar topologia e paths por unidade.
- Inclusão do fix pack `ECG-BDE-Fix.ps1` com launcher dedicado `ECG-BDE-Fix_Menu.bat`.
- Inclusão do runbook operacional de correção para cenários ECGv6/BDE e NETDIR.

### Changed
- Evolução do core `ELCE_ECG_Diagnostics.ps1` para resolução orientada a perfis de unidade.
- Execução via wrappers BAT/CMD com `-ExecutionPolicy Bypass`, reduzindo intervenção manual.
- Organização do repositório ampliada com `fixpacks/` e `docs/runbooks/`.

### Notes
- O core principal permanece read-only.
- A remediação continua separada do fluxo de laudo por governança de produto.

## [1.0.1] - 2026-03-26

### Fixed
- Corrigida a persistência dos artefatos humanos (`ELCE_ECG_Diagnostics_Report.html` e `ELCE_ECG_Diagnostics_Summary.txt`) para UTF-8 com BOM, reduzindo problemas de acentuação em abertura no Windows.
- Corrigida a serialização do gráfico temporal para propriedades compatíveis com o JavaScript embutido (`labels/series/name/values/max/color`), eliminando o mismatch de case que impedia a renderização.
- Ampliada a classificação de tipo de máquina com heurísticas por hostname e fallback por sistema operacional/artefatos locais.

### Changed
- Adicionada exposição da origem da classificação e do sistema operacional no HTML técnico.
- Lineage interna do core promovida para `3.2-html-first-hotfix3`.

## [1.0.0] - 2026-03-26

### Added
- Estrutura inicial do repositório GitHub
- Core baseline em `src/`
- Documentação técnica em `docs/`
- Templates de Issue e Pull Request
- Workflow leve para empacotamento de release

### Notes
- Esta release representa a baseline operacional governada.
- Lineage interna do core: `3.2-html-first-hotfix2`
