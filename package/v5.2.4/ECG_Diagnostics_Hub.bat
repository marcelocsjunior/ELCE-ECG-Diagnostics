@echo off
setlocal EnableExtensions DisableDelayedExpansion
title ECG Diagnostics Hub v5.2.4

set "SCRIPT_DIR=%~dp0"
set "CORE_PS1=%SCRIPT_DIR%ECG_Diagnostics_Core.ps1"
set "PROFILE_INI=%SCRIPT_DIR%ECG_FieldKit.ini"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

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

if not exist "%PROFILE_INI%" (
  echo [ERRO] Arquivo ausente:
  echo %PROFILE_INI%
  echo.
  pause
  exit /b 1
)

:MENU
cls
echo =====================================================
echo     ECG Diagnostics Hub - Menu Unificado v5.2.4
echo =====================================================
echo.
echo 1^) Correcao completa ^(Fix^) - corrige NETDIR, IDAPI32.CFG
echo 2^) Diagnostico somente leitura ^(Auto^) - sem alteracoes
echo 3^) Comparar dois laudos/estacoes ^(Compare^)
echo 4^) Rollback de registro ^(.reg^)
echo 5^) Monitoramento de desempenho ^(usa MonitorMinutes/SampleInterval do INI^)
echo 6^) Coleta estatica de informacoes ^(JSON^)
echo.
echo 0^) Sair
echo.
set /p OPT=Escolha uma opcao [0-6]: 

if "%OPT%"=="1" goto FIX
if "%OPT%"=="2" goto AUTO
if "%OPT%"=="3" goto COMPARE
if "%OPT%"=="4" goto ROLLBACK
if "%OPT%"=="5" goto MONITOR
if "%OPT%"=="6" goto STATIC
if "%OPT%"=="0" exit /b 0

echo.
echo [ERRO] Opcao invalida.
pause
goto MENU

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
set "RUN_ARGS=-Mode Fix -ProfilePath ""%PROFILE_INI%"""
goto RUNCORE

:AUTO
set "RUN_ARGS=-Mode Auto -ProfilePath ""%PROFILE_INI%"""
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
  set "RUN_ARGS=-Mode Compare -ProfilePath ""%PROFILE_INI%"""
  goto RUNCORE
)

if "%COMPARE_LEFT%"=="" goto COMPARE_PATH_ERROR
if "%COMPARE_RIGHT%"=="" goto COMPARE_PATH_ERROR

set "RUN_ARGS=-Mode Compare -ProfilePath ""%PROFILE_INI%"" -CompareLeftReport ""%COMPARE_LEFT%"" -CompareRightReport ""%COMPARE_RIGHT%"""
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
set "RUN_ARGS=-Mode Rollback -ProfilePath ""%PROFILE_INI%"" -RollbackFile ""%ROLLBACK_PATH%"""
goto RUNCORE

:MONITOR
set "RUN_ARGS=-Mode Monitor -ProfilePath ""%PROFILE_INI%"""
goto RUNCORE

:STATIC
set "RUN_ARGS=-Mode CollectStatic -ProfilePath ""%PROFILE_INI%"""
goto RUNCORE
