ELCE ECG - Pacote unificado final

Conteúdo:
- fixpacks/BDE-Fix-Core.ps1       -> core único
- fixpacks/ECG-BDE-Fix_Menu.bat   -> menu único
- fixpacks/ECG-BDE-Fix.ps1        -> shim legado
- fixpacks/ECG_UnitProfiles.json  -> profiles válidos
- fixpacks/runbookECG-BDE-Fix.txt -> runbook
- src/ELCE_ECG_Diagnostics.ps1    -> core oficial do diagnóstico (gera HTML)

Observações:
- Opção 1 do menu chama o diagnóstico oficial e gera o HTML em:
  C:\ECG\Output\Latest\ELCE_ECG_Diagnostics_Report.html
- Opção 8 abre o último laudo HTML
- Opção 5 executa todas as correções BDE
