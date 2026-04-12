@echo off
setlocal EnableExtensions DisableDelayedExpansion
title ECG Diagnostics Hub v6.3.2

set "SCRIPT_DIR=%~dp0"
set "CORE_PS1=%SCRIPT_DIR%ECG_Diagnostics_Core_v6_3_2.ps1"
set "TEMPLATE_INI=%SCRIPT_DIR%ECG_FieldKit_Unified_v6_3_2.ini"
set "BUILDER_PS1=%SCRIPT_DIR%ECG_ProfileBuilder_v6_3_2.ps1"
set "AI_HUB=%SCRIPT_DIR%ECG_AI_Prototype_Hub_v0_2_1.bat"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "RUNTIME_INI=%TEMP%\ECG_FieldKit_Runtime_v6_3_2.ini"

if not exist "%PS_EXE%" (
  echo [ERRO] powershell.exe nao encontrado em:
  echo %PS_EXE%
  echo.
  pause
  exit /b 1
)
if not exist "%CORE_PS1%" (
  echo [ERRO] Arquivo ausente:
  echo %CORE_PS1%
  echo.
  pause
  exit /b 1
)
if not exist "%TEMPLATE_INI%" (
  echo [ERRO] Arquivo ausente:
  echo %TEMPLATE_INI%
  echo.
  pause
  exit /b 1
)
if not exist "%BUILDER_PS1%" (
  echo [ERRO] Arquivo ausente:
  echo %BUILDER_PS1%
  echo.
  pause
  exit /b 1
)

:MENU
cls
echo ======================================================================
echo   ECG Diagnostics Hub v6.3.2 - hardening de target + custom + duracao
echo ======================================================================
echo.
echo 1^) Correcao autorizada ^(Fix^) - escolhe 1 target e duracao
echo 2^) Detectar + decidir + relatar ^(Detect^) - escolhe 1 target e duracao
echo 3^) Diagnostico somente leitura legado ^(Auto^) - escolhe 1 target e duracao
echo 4^) Comparar dois laudos legados ^(Compare / ECG_Report.json^)
echo 5^) Rollback de registro ^(.reg^)
echo 6^) Monitoramento de desempenho ^(Monitor^) - escolhe 1 target e duracao
echo 7^) Coleta estatica de informacoes ^(CollectStatic^)
echo 8^) Diagnostico single-target ^(Single^) - escolhe 1 target e duracao
echo 9^) CompareBackend guiado ^(presets/custom + tempo^)
echo J^) Comparar dois JSONs da suite unificada ^(CompareJson^)
echo A^) Abrir pasta do pacote
echo B^) Abrir INI base
echo I^) Abrir Hub IA local
echo.
echo 0^) Sair
echo.
set /p OPT=Escolha uma opcao [0-9,J,A,B,I]: 

if /I "%OPT%"=="1" goto FIX
if /I "%OPT%"=="2" goto DETECT
if /I "%OPT%"=="3" goto AUTO
if /I "%OPT%"=="4" goto COMPARE
if /I "%OPT%"=="5" goto ROLLBACK
if /I "%OPT%"=="6" goto MONITOR
if /I "%OPT%"=="7" goto STATIC
if /I "%OPT%"=="8" goto SINGLE
if /I "%OPT%"=="9" goto COMPAREBACKEND
if /I "%OPT%"=="J" goto COMPAREJSON
if /I "%OPT%"=="A" goto OPENFOLDER
if /I "%OPT%"=="B" goto OPENINI
if /I "%OPT%"=="I" goto OPENAIHUB
if "%OPT%"=="0" exit /b 0

echo.
echo [ERRO] Opcao invalida.
pause
goto MENU

:RESET_TARGET_STATE
set "TARGET_KEY="
set "PRIMARY_TARGET="
set "SELECTED_TARGETS="
set "CUSTOM_LABEL="
set "CUSTOM_DB="
set "CUSTOM_NETDIR="
set "CUSTOM_HOSTOS="
set "CUSTOM_SMB="
set "CUSTOM_LEGACY="
set "TEST_MINUTES="
set "TEST_INTERVAL="
exit /b 0

:ASK_PRIMARY_TARGET
call :RESET_TARGET_STATE
echo.
echo Escolha o target principal:
echo   1^) FS01 legado por hostname ^(\\SRVVM1-FS01\FS\ECG\HW\DATABASE^)
echo   2^) XP legado por IP ^(\\192.168.1.57\Database^)
echo   3^) WS2016 novo oficial ^(\\SRVVM1-ECG\DBE^)
echo   4^) Custom ^(informar DB e NetDir^)
set /p TARGET_OPT=Target [1-4]: 
if "%TARGET_OPT%"=="1" set "TARGET_KEY=FS01"
if "%TARGET_OPT%"=="2" set "TARGET_KEY=XP"
if "%TARGET_OPT%"=="3" set "TARGET_KEY=WS2016"
if "%TARGET_OPT%"=="4" set "TARGET_KEY=CUSTOM"
if "%TARGET_KEY%"=="" (
  echo [ERRO] Target invalido.
  pause
  exit /b 1
)
if /I "%TARGET_KEY%"=="CUSTOM" (
  call :ASK_CUSTOM_TARGET
  if errorlevel 1 exit /b 1
)
exit /b 0

:ASK_CUSTOM_TARGET
set /p CUSTOM_LABEL=Label do target custom [CUSTOM_TARGET]: 
if "%CUSTOM_LABEL%"=="" set "CUSTOM_LABEL=CUSTOM_TARGET"
set /p CUSTOM_DB=Caminho UNC do DB: 
if "%CUSTOM_DB%"=="" (
  echo [ERRO] Caminho do DB nao informado.
  pause
  exit /b 1
)
if not "%CUSTOM_DB:~0,2%"=="\\" (
  echo [ERRO] Caminho do DB deve ser UNC iniciado por \\\.
  pause
  exit /b 1
)
set /p CUSTOM_NETDIR=Caminho UNC do NetDir: 
if "%CUSTOM_NETDIR%"=="" (
  echo [ERRO] Caminho do NetDir nao informado.
  pause
  exit /b 1
)
if not "%CUSTOM_NETDIR:~0,2%"=="\\" (
  echo [ERRO] Caminho do NetDir deve ser UNC iniciado por \\\.
  pause
  exit /b 1
)
set /p CUSTOM_HOSTOS=Host OS hint [Custom target]: 
if "%CUSTOM_HOSTOS%"=="" set "CUSTOM_HOSTOS=Custom target"
set /p CUSTOM_SMB=SMB hint [SMB custom]: 
if "%CUSTOM_SMB%"=="" set "CUSTOM_SMB=SMB custom"
set /p CUSTOM_LEGACY=LegacyHint? [S/N]: 
if "%CUSTOM_LEGACY%"=="" set "CUSTOM_LEGACY=N"
exit /b 0

:ASK_TIME_OPERATIONAL
echo.
set /p TEST_MINUTES=Tempo da coleta em minutos [3]: 
if "%TEST_MINUTES%"=="" set "TEST_MINUTES=3"
call :ENSURE_POSITIVE_INT TEST_MINUTES "Tempo da coleta"
if errorlevel 1 exit /b 1
set /p TEST_INTERVAL=Intervalo entre amostras em segundos [15]: 
if "%TEST_INTERVAL%"=="" set "TEST_INTERVAL=15"
call :ENSURE_POSITIVE_INT TEST_INTERVAL "Intervalo entre amostras"
if errorlevel 1 exit /b 1
exit /b 0

:ASK_TIME_COMPARE
echo.
set /p TEST_MINUTES=Tempo da coleta em minutos [3]: 
if "%TEST_MINUTES%"=="" set "TEST_MINUTES=3"
call :ENSURE_POSITIVE_INT TEST_MINUTES "Tempo da coleta"
if errorlevel 1 exit /b 1
set /p TEST_INTERVAL=Intervalo entre amostras em segundos [3]: 
if "%TEST_INTERVAL%"=="" set "TEST_INTERVAL=3"
call :ENSURE_POSITIVE_INT TEST_INTERVAL "Intervalo entre amostras"
if errorlevel 1 exit /b 1
exit /b 0

:ASK_COMPARE_TARGETS
call :RESET_TARGET_STATE
echo.
echo Escolha os targets do CompareBackend ^(separados por virgula^):
echo   1^) FS01 legado por hostname
echo   2^) XP legado por IP
echo   3^) WS2016 novo oficial
echo   4^) Custom
set /p COMPARE_SELECT=Selecao [ex.: 2,3 ou 1,2,3,4]: 
if "%COMPARE_SELECT%"=="" (
  echo [ERRO] Informe pelo menos 1 target. Se informar 1 so, o target principal diferente sera adicionado automaticamente.
  pause
  exit /b 1
)
call :MAP_COMPARE_SELECTION "%COMPARE_SELECT%"
if errorlevel 1 exit /b 1
if /I "%SELECTED_TARGETS%"=="CUSTOM" (
  echo [ERRO] CompareBackend requer 2 targets distintos ou 1 target + principal diferente.
  pause
  exit /b 1
)
set /p PRIMARY_TARGET=Qual target sera o principal do perfil [FS01/XP/WS2016/CUSTOM] [WS2016]: 
if "%PRIMARY_TARGET%"=="" set "PRIMARY_TARGET=WS2016"
call :VALIDATE_TARGET_KEY PRIMARY_TARGET
if errorlevel 1 exit /b 1
exit /b 0

:MAP_COMPARE_SELECTION
set "RAWSEL=%~1"
set "TMPSEL=%RAWSEL:,= %"
set "SELECTED_TARGETS="
for %%G in (%TMPSEL%) do (
  if "%%~G"=="1" call :ADD_TARGET FS01
  if "%%~G"=="2" call :ADD_TARGET XP
  if "%%~G"=="3" call :ADD_TARGET WS2016
  if "%%~G"=="4" call :ADD_TARGET CUSTOM
)
if "%SELECTED_TARGETS%"=="" (
  echo [ERRO] Nenhum target valido selecionado.
  pause
  exit /b 1
)
if not "%SELECTED_TARGETS:CUSTOM=%"=="%SELECTED_TARGETS%" (
  call :ASK_CUSTOM_TARGET
  if errorlevel 1 exit /b 1
)
exit /b 0

:ADD_TARGET
if "%SELECTED_TARGETS%"=="" (
  set "SELECTED_TARGETS=%~1"
) else (
  echo %SELECTED_TARGETS% | find /I "%~1" >nul
  if errorlevel 1 set "SELECTED_TARGETS=%SELECTED_TARGETS%,%~1"
)
exit /b 0

:BUILD_OPERATIONAL_PROFILE
set "BUILD_SCENARIO=Operational"
set "BUILD_PRIMARY=%TARGET_KEY%"
set "BUILD_SELECTED=%TARGET_KEY%"
goto RUN_BUILD_PROFILE

:BUILD_SINGLE_PROFILE
set "BUILD_SCENARIO=Single"
set "BUILD_PRIMARY=%TARGET_KEY%"
set "BUILD_SELECTED=%TARGET_KEY%"
goto RUN_BUILD_PROFILE

:BUILD_COMPARE_PROFILE
set "BUILD_SCENARIO=CompareBackend"
set "BUILD_PRIMARY=%PRIMARY_TARGET%"
set "BUILD_SELECTED=%SELECTED_TARGETS%"
goto RUN_BUILD_PROFILE

:RUN_BUILD_PROFILE
if not "%BUILD_SELECTED:CUSTOM=%"=="%BUILD_SELECTED%" (
  if /I "%CUSTOM_LEGACY%"=="S" (
    "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario "%BUILD_SCENARIO%" -PrimaryTargetKey "%BUILD_PRIMARY%" -SelectedTargetsCsv "%BUILD_SELECTED%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL% -CustomLabel "%CUSTOM_LABEL%" -CustomDbPath "%CUSTOM_DB%" -CustomNetDirPath "%CUSTOM_NETDIR%" -CustomHostOsHint "%CUSTOM_HOSTOS%" -CustomSmbDialectHint "%CUSTOM_SMB%" -CustomLegacyHint
  ) else (
    "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario "%BUILD_SCENARIO%" -PrimaryTargetKey "%BUILD_PRIMARY%" -SelectedTargetsCsv "%BUILD_SELECTED%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL% -CustomLabel "%CUSTOM_LABEL%" -CustomDbPath "%CUSTOM_DB%" -CustomNetDirPath "%CUSTOM_NETDIR%" -CustomHostOsHint "%CUSTOM_HOSTOS%" -CustomSmbDialectHint "%CUSTOM_SMB%"
  )
) else (
  "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario "%BUILD_SCENARIO%" -PrimaryTargetKey "%BUILD_PRIMARY%" -SelectedTargetsCsv "%BUILD_SELECTED%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL%
)
set "BRC=%ERRORLEVEL%"
if not "%BRC%"=="0" (
  echo.
  echo [ERRO] Falha ao gerar perfil runtime.
  pause
  exit /b 1
)
if not exist "%RUNTIME_INI%" (
  echo.
  echo [ERRO] Perfil runtime nao foi gerado.
  pause
  exit /b 1
)
echo [OK] Perfil runtime: %RUNTIME_INI%
exit /b 0

:RUNCORE
echo.
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CORE_PS1%" %RUN_ARGS%
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
  echo [OK] Execucao concluida com sucesso.
) else (
  echo [ERRO] Execucao retornou codigo %RC%.
)
echo.
pause
goto MENU

:FIX
call :ASK_PRIMARY_TARGET
if errorlevel 1 goto MENU
call :ASK_TIME_OPERATIONAL
if errorlevel 1 goto MENU
call :BUILD_OPERATIONAL_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode Fix -AuthorizedRemediation -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:DETECT
call :ASK_PRIMARY_TARGET
if errorlevel 1 goto MENU
call :ASK_TIME_OPERATIONAL
if errorlevel 1 goto MENU
call :BUILD_OPERATIONAL_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode Detect -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:AUTO
call :ASK_PRIMARY_TARGET
if errorlevel 1 goto MENU
call :ASK_TIME_OPERATIONAL
if errorlevel 1 goto MENU
call :BUILD_OPERATIONAL_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode Auto -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:COMPARE
echo.
echo Informe os caminhos completos dos dois arquivos ECG_Report.json.
echo Deixe ambos vazios para comparar automaticamente os 2 laudos mais recentes do OutDir.
echo.
set "COMPARE_LEFT="
set "COMPARE_RIGHT="
set /p COMPARE_LEFT=Relatorio esquerdo (.json): 
set /p COMPARE_RIGHT=Relatorio direito (.json): 
if "%COMPARE_LEFT%"=="" if "%COMPARE_RIGHT%"=="" (
  set "RUN_ARGS=-Mode Compare -ProfilePath ""%TEMPLATE_INI%"" -OpenReport"
  goto RUNCORE
)
if "%COMPARE_LEFT%"=="" goto COMPARE_PATH_ERROR
if "%COMPARE_RIGHT%"=="" goto COMPARE_PATH_ERROR
set "RUN_ARGS=-Mode Compare -ProfilePath ""%TEMPLATE_INI%"" -CompareLeftReport ""%COMPARE_LEFT%"" -CompareRightReport ""%COMPARE_RIGHT%"" -OpenReport"
goto RUNCORE

:COMPARE_PATH_ERROR
echo.
echo [ERRO] Informe os dois caminhos de laudo JSON ou deixe ambos vazios.
pause
goto MENU

:ROLLBACK
echo.
set /p ROLLBACK_PATH=Informe o caminho completo do arquivo .reg: 
if "%ROLLBACK_PATH%"=="" (
  echo [ERRO] Caminho nao informado.
  echo.
  pause
  goto MENU
)
set "RUN_ARGS=-Mode Rollback -ProfilePath ""%TEMPLATE_INI%"" -RollbackFile ""%ROLLBACK_PATH%"""
goto RUNCORE

:MONITOR
call :ASK_PRIMARY_TARGET
if errorlevel 1 goto MENU
call :ASK_TIME_OPERATIONAL
if errorlevel 1 goto MENU
call :BUILD_OPERATIONAL_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode Monitor -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:STATIC
set "RUN_ARGS=-Mode CollectStatic -ProfilePath ""%TEMPLATE_INI%"""
goto RUNCORE

:SINGLE
call :ASK_PRIMARY_TARGET
if errorlevel 1 goto MENU
call :ASK_TIME_COMPARE
if errorlevel 1 goto MENU
call :BUILD_SINGLE_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode Single -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:COMPAREBACKEND
call :ASK_COMPARE_TARGETS
if errorlevel 1 goto MENU
call :ASK_TIME_COMPARE
if errorlevel 1 goto MENU
call :BUILD_COMPARE_PROFILE
if errorlevel 1 goto MENU
set "RUN_ARGS=-Mode CompareBackend -ProfilePath ""%RUNTIME_INI%"" -OpenReport"
goto RUNCORE

:COMPAREJSON
echo.
echo Informe os caminhos completos dos dois arquivos JSON gerados pela suite unificada.
echo.
set "COMPARE_LEFT="
set "COMPARE_RIGHT="
set /p COMPARE_LEFT=Relatorio esquerdo (.json): 
set /p COMPARE_RIGHT=Relatorio direito (.json): 
if "%COMPARE_LEFT%"=="" goto COMPAREJSON_ERROR
if "%COMPARE_RIGHT%"=="" goto COMPAREJSON_ERROR
set "RUN_ARGS=-Mode CompareJson -ProfilePath ""%TEMPLATE_INI%"" -CompareLeftReport ""%COMPARE_LEFT%"" -CompareRightReport ""%COMPARE_RIGHT%"" -OpenReport"
goto RUNCORE

:COMPAREJSON_ERROR
echo.
echo [ERRO] Informe os dois caminhos dos relatorios JSON.
pause
goto MENU

:OPENFOLDER
start "" "%SCRIPT_DIR%"
goto MENU

:OPENINI
start "" notepad.exe "%TEMPLATE_INI%"
goto MENU

:OPENAIHUB
if exist "%AI_HUB%" (
  start "" "%AI_HUB%"
) else (
  echo [ERRO] Hub IA nao encontrado.
  pause
)
goto MENU

:VALIDATE_TARGET_KEY
call set "_CURRENT_VALUE=%%%~1%%"
if /I "%_CURRENT_VALUE%"=="FS01" exit /b 0
if /I "%_CURRENT_VALUE%"=="XP" exit /b 0
if /I "%_CURRENT_VALUE%"=="WS2016" exit /b 0
if /I "%_CURRENT_VALUE%"=="CUSTOM" exit /b 0
echo [ERRO] Target invalido: %_CURRENT_VALUE%
pause
exit /b 1

:ENSURE_POSITIVE_INT
call set "_NUMVALUE=%%%~1%%"
echo(%_NUMVALUE%| findstr /r "^[1-9][0-9]*$" >nul
if errorlevel 1 (
  echo [ERRO] %~2 invalido: %_NUMVALUE%
  pause
  exit /b 1
)
exit /b 0
