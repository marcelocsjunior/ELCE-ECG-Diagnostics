ELCE ECG - referência de pacote unificado

O repositório passa a trabalhar com dois artefatos de release:

1) SOURCE PACKAGE
- snapshot do repositório
- preserva a árvore versionada (`src/`, `docs/`, `fixpacks/`, etc.)
- indicado para auditoria, revisão técnica e uso direto do clone

2) DEPLOY PACKAGE
- entrega layout operacional pronto para cópia em `C:\ECG\Tool`
- mantém o diagnóstico e a remediação em trilhas separadas, porém operacionalmente acessíveis

Conteúdo mínimo esperado no deploy package:

Tool/ELCE_ECG_Diagnostics.ps1
Tool/ELCE_ECG_Diagnostics_Menu.bat
Tool/ExecutarDiagnostico.cmd
Tool/ECG_UnitProfiles.json
Tool/ECG-BDE-Fix.ps1
Tool/ECG-BDE-Fix_Menu.bat
Tool/BDE-Fix-Core.ps1
Tool/runbookECG-BDE-Fix.txt
docs/runbooks/runbookECG-BDE-Fix.txt
README-DEPLOY.txt

Observações:
- O runbook canônico do produto permanece em `docs/runbooks/runbookECG-BDE-Fix.txt`
- A cópia em `Tool/runbookECG-BDE-Fix.txt` existe por compatibilidade operacional do menu
- A opção 1 do menu chama o diagnóstico oficial e gera o HTML em:
  C:\ECG\Output\Latest\ELCE_ECG_Diagnostics_Report.html
- A opção 8 do fix menu abre o último laudo HTML
- A opção 5 do fix menu executa todas as correções BDE
