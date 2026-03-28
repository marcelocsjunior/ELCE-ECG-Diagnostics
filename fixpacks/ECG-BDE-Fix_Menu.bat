@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "TOOL_DIR=%~dp0"
set "SCRIPT=%TOOL_DIR%ECG-BDE-Fix.ps1"
set "RUNBOOK=%TOOL_DIR%runbookECG-BDE-Fix.txt"
set "PROFILE=%TOOL_DIR%ECG_UnitProfiles.json"

if not exist "%SCRIPT%" (
    echo.
    echo [ERRO] Script de correcao nao encontrado:
    echo %SCRIPT%
    echo.
    pause
    exit /b 1
)

:MENU
cls
echo ============================================
echo   ECG BDE Fix
echo ============================================
echo.
echo [1] Diagnosticar sem alterar
echo [2] Aplicar correcao do NETDIR
echo [3] Aplicar correcao + criar diretorios locais padrao
echo [4] Abrir runbook
echo [0] Sair
echo.
set /p opt="Escolha uma opcao: "

if "%opt%"=="1" goto DIAG
if "%opt%"=="2" goto FIX
if "%opt%"=="3" goto FIXDIR
if "%opt%"=="4" goto RUNBOOK
if "%opt%"=="0" exit /b 0
goto MENU

:DIAG
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Unit AUTO -ProfileFile "%PROFILE%"
pause
goto MENU

:FIX
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Fix -Unit AUTO -ProfileFile "%PROFILE%"
pause
goto MENU

:FIXDIR
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -Fix -CreateMissingDirs -Unit AUTO -ProfileFile "%PROFILE%"
pause
goto MENU

:RUNBOOK
if exist "%RUNBOOK%" (
    start "" "%RUNBOOK%"
) else (
    echo.
    echo Runbook nao encontrado:
    echo %RUNBOOK%
    echo.
    pause
)
goto MENU
