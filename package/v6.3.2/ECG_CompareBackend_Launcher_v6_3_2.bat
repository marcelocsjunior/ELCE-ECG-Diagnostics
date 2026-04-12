@echo off
setlocal EnableExtensions DisableDelayedExpansion
title ECG CompareBackend Launcher v6.3.2

set "SCRIPT_DIR=%~dp0"
set "CORE_PS1=%SCRIPT_DIR%ECG_Diagnostics_Core_v6_3_2.ps1"
set "TEMPLATE_INI=%SCRIPT_DIR%ECG_FieldKit_Unified_v6_3_2.ini"
set "BUILDER_PS1=%SCRIPT_DIR%ECG_ProfileBuilder_v6_3_2.ps1"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "RUNTIME_INI=%TEMP%\ECG_FieldKit_Runtime_Compare_v6_3_2.ini"

if not exist "%PS_EXE%" (
  echo [ERRO] powershell.exe nao encontrado.
  pause
  exit /b 1
)
if not exist "%CORE_PS1%" (
  echo [ERRO] Core nao encontrado:
  echo %CORE_PS1%
  pause
  exit /b 1
)
if not exist "%TEMPLATE_INI%" (
  echo [ERRO] INI nao encontrado:
  echo %TEMPLATE_INI%
  pause
  exit /b 1
)
if not exist "%BUILDER_PS1%" (
  echo [ERRO] Builder nao encontrado:
  echo %BUILDER_PS1%
  pause
  exit /b 1
)

echo ======================================================================
echo ECG CompareBackend Launcher v6.3.2
echo ======================================================================
echo.
echo Targets disponiveis:
echo   1^) FS01 legado por hostname
echo   2^) XP legado por IP
echo   3^) WS2016 novo oficial
echo   4^) Custom
echo.
set /p COMPARE_SELECT=Selecao [ex.: 2,3 ou 1,2,3,4]: 
if "%COMPARE_SELECT%"=="" (
  echo [ERRO] Informe pelo menos 1 target. Se informar 1 so, o target principal diferente sera adicionado automaticamente.
  pause
  exit /b 1
)

call :MAP_COMPARE_SELECTION "%COMPARE_SELECT%"
if errorlevel 1 exit /b 1
if "%SELECTED_TARGETS%"=="" (
  echo [ERRO] Nenhum target valido selecionado.
  pause
  exit /b 1
)
if /I "%SELECTED_TARGETS%"=="CUSTOM" (
  echo [ERRO] CompareBackend requer 2 targets distintos ou 1 target + principal diferente.
  pause
  exit /b 1
)

set /p PRIMARY_TARGET=Target principal [FS01/XP/WS2016/CUSTOM] [WS2016]: 
if "%PRIMARY_TARGET%"=="" set "PRIMARY_TARGET=WS2016"
call :VALIDATE_TARGET_KEY PRIMARY_TARGET
if errorlevel 1 exit /b 1

set /p TEST_MINUTES=Tempo da coleta em minutos [3]: 
if "%TEST_MINUTES%"=="" set "TEST_MINUTES=3"
call :ENSURE_POSITIVE_INT TEST_MINUTES "Tempo da coleta"
if errorlevel 1 exit /b 1

set /p TEST_INTERVAL=Intervalo entre amostras em segundos [3]: 
if "%TEST_INTERVAL%"=="" set "TEST_INTERVAL=3"
call :ENSURE_POSITIVE_INT TEST_INTERVAL "Intervalo entre amostras"
if errorlevel 1 exit /b 1

echo.
call :RUN_BUILDER
if errorlevel 1 exit /b 1

echo.
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%CORE_PS1%" -Mode CompareBackend -ProfilePath "%RUNTIME_INI%" -OpenReport
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
  echo [OK] CompareBackend concluido.
) else (
  echo [ERRO] CompareBackend retornou codigo %RC%.
)
echo.
pause
exit /b %RC%

:RUN_BUILDER
if not "%SELECTED_TARGETS:CUSTOM=%"=="%SELECTED_TARGETS%" (
  if /I "%CUSTOM_LEGACY%"=="S" (
    "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario CompareBackend -PrimaryTargetKey "%PRIMARY_TARGET%" -SelectedTargetsCsv "%SELECTED_TARGETS%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL% -CustomLabel "%CUSTOM_LABEL%" -CustomDbPath "%CUSTOM_DB%" -CustomNetDirPath "%CUSTOM_NETDIR%" -CustomHostOsHint "%CUSTOM_HOSTOS%" -CustomSmbDialectHint "%CUSTOM_SMB%" -CustomLegacyHint
  ) else (
    "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario CompareBackend -PrimaryTargetKey "%PRIMARY_TARGET%" -SelectedTargetsCsv "%SELECTED_TARGETS%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL% -CustomLabel "%CUSTOM_LABEL%" -CustomDbPath "%CUSTOM_DB%" -CustomNetDirPath "%CUSTOM_NETDIR%" -CustomHostOsHint "%CUSTOM_HOSTOS%" -CustomSmbDialectHint "%CUSTOM_SMB%"
  )
) else (
  "%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%BUILDER_PS1%" -TemplatePath "%TEMPLATE_INI%" -OutputPath "%RUNTIME_INI%" -Scenario CompareBackend -PrimaryTargetKey "%PRIMARY_TARGET%" -SelectedTargetsCsv "%SELECTED_TARGETS%" -Minutes %TEST_MINUTES% -IntervalSeconds %TEST_INTERVAL%
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

:ASK_CUSTOM_TARGET
set "CUSTOM_LABEL="
set "CUSTOM_DB="
set "CUSTOM_NETDIR="
set "CUSTOM_HOSTOS="
set "CUSTOM_SMB="
set "CUSTOM_LEGACY="
set /p CUSTOM_LABEL=Label do target custom [CUSTOM_TARGET]: 
if "%CUSTOM_LABEL%"=="" set "CUSTOM_LABEL=CUSTOM_TARGET"
set /p CUSTOM_DB=Caminho UNC do DB: 
if "%CUSTOM_DB%"=="" (
  echo [ERRO] Caminho do DB nao informado.
  pause
  exit /b 1
)
if not "%CUSTOM_DB:~0,2%"=="\\" (
  echo [ERRO] DB custom deve ser caminho UNC iniciado por \\\.
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
  echo [ERRO] NetDir custom deve ser caminho UNC iniciado por \\\.
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

:VALIDATE_TARGET_KEY
call set "_CURRENT_VALUE=%%%~1%%"
if /I "%_CURRENT_VALUE%"=="FS01" exit /b 0
if /I "%_CURRENT_VALUE%"=="XP" exit /b 0
if /I "%_CURRENT_VALUE%"=="WS2016" exit /b 0
if /I "%_CURRENT_VALUE%"=="CUSTOM" exit /b 0
echo [ERRO] Target principal invalido: %_CURRENT_VALUE%
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
