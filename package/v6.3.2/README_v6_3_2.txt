ECG Diagnostics Suite v6.3.2 - Hardening Release

Escopo desta release
- Consolidar o backend oficial WS2016 dedicado
- Manter FS01 e XP como alvos de comparacao A/B ou contingencia
- Endurecer o pacote para operacao via menu e execucao direta
- Preservar fluxo: Detectar -> Decidir -> Relatar -> Corrigir somente se autorizado

Arquivos principais
- ECG_Diagnostics_Core_v6_3_2.ps1
- ECG_Diagnostics_Hub_v6_3_2.bat
- ECG_CompareBackend_Launcher_v6_3_2.bat
- ECG_ProfileBuilder_v6_3_2.ps1
- ECG_FieldKit_Unified_v6_3_2.ini
- ECG_AI_Prototype_Hub_v0_2_1.bat
- ECG_Diagnostics_AI_Prototype_v0_2_1.ps1

Presets oficiais de target
1. FS01 legado por hostname
   - DB     = \\SRVVM1-FS01\FS\ECG\HW\DATABASE
   - NetDir = \\SRVVM1-FS01\FS\ECG\HW\DATABASE\NetDir
2. XP legado por IP
   - DB     = \\192.168.1.57\Database
   - NetDir = \\192.168.1.57\Database\NetDir
3. WS2016 novo oficial
   - DB     = \\SRVVM1-ECG\DBE
   - NetDir = \\SRVVM1-ECG\DBE\NetDir
4. Custom
   - operador informa Label, DB e NetDir

Hardening aplicado
- Fonte de verdade dos presets movida para o INI base via PresetKey
- Builder passou a montar runtime lendo os alvos do INI, sem hardcode paralelo
- Core passou a resolver perfil padrao priorizando v6.3.2 e v6.3.1 antes dos perfis legados
- Parametro -TargetLabel formalizado no core para Single mode
- Export do rollback .reg ajustado para Unicode BOM, mais seguro para reg.exe/regedit
- Launcher CompareBackend endurecido para validar target principal, minutos, intervalo e custom incompleto
- Hub endurecido com as mesmas validacoes e melhor propagacao de erro
- Hub IA local passou a apontar para o perfil v6.3.2

Observacao honesta
- Validacao feita com leitura completa, parser estatico PowerShell, checagem de consistencia e empacotamento.
- Homologacao final continua obrigatoria em Windows real com PowerShell 5.1 antes de rollout amplo.
