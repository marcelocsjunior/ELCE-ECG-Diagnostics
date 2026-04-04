@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul 2>&1
color 0A
title ECGv6 FieldKit

set "SCRIPT_DIR=%~dp0"
if "%SCRIPT_DIR%"=="" set "SCRIPT_DIR=.\"
set "PS1=%SCRIPT_DIR%ECGv6_FieldKit.ps1"
set "INI=%SCRIPT_DIR%ECGv6_FieldKit.ini"
set "OUTDIR=C:\ECG\FieldKit"

if not exist "%PS1%" (
  echo.
  echo [ERRO] Nao encontrei o motor principal:
  echo %PS1%
  echo.
  pause
  exit /b 1
)

call :refresh_config

:menu
call :refresh_config
cls
echo ================================================================
echo   ECGv6 FieldKit - Menu Operacional
echo   Diagnostico, correcao, comparacao e rollback controlado
echo ================================================================
echo.
echo Script  : %PS1%
echo INI     : %INI%
echo Saida   : %OUTDIR%
echo Host    : %COMPUTERNAME%
echo.
echo [1] Prepare ^(cria/verifica diretorios esperados^)
echo [2] Audit ^(somente laudo^)
echo [3] Auto ^(laudo + write probe + abre HTML^)
echo [4] Fix ^(auto-fix seguro + write probe + abre HTML^)
echo [5] Compare ^(compara os 2 laudos JSON mais recentes^)
echo [6] Editar INI
echo [7] Abrir pasta de saida
echo [8] Rollback por arquivo .reg
echo [9] Sair
echo.
set "OPT="
set /p OPT=Escolha uma opcao [1-9]: 

if "%OPT%"=="1" goto prepare
if "%OPT%"=="2" goto audit
if "%OPT%"=="3" goto auto
if "%OPT%"=="4" goto fix
if "%OPT%"=="5" goto compare
if "%OPT%"=="6" goto editini
if "%OPT%"=="7" goto openout
if "%OPT%"=="8" goto rollback
if "%OPT%"=="9" exit /b 0

echo.
echo [WARN] Opcao invalida.
timeout /t 1 >nul 2>&1
goto menu

:refresh_config
for /f "usebackq delims=" %%I in (`powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -Command "$p='%INI%'; if(Test-Path -LiteralPath $p){ $line = Get-Content -LiteralPath $p ^| Where-Object { $_ -match '^[ ]*OutDir[ ]*=' } ^| Select-Object -First 1; if($line){ ($line -split '=',2)[1].Trim() } else { 'C:\ECG\FieldKit' } } else { 'C:\ECG\FieldKit' }"`) do set "OUTDIR=%%I"
if not defined OUTDIR set "OUTDIR=C:\ECG\FieldKit"
if not exist "%OUTDIR%" mkdir "%OUTDIR%" >nul 2>&1
exit /b 0

:run_common
echo.
echo [INFO] Executando modo %~1 ...
echo.
powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%" -Mode %~1 -ProfilePath "%INI%" -OutDir "%OUTDIR%" %~2 %~3 %~4 %~5 %~6 %~7 %~8 %~9
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
  echo [OK] Execucao concluida.
) else (
  echo [ERRO] PowerShell retornou codigo %RC%.
)
echo.
pause
goto menu

:prepare
call :run_common Prepare -OpenReport

:audit
call :run_common Audit -OpenReport

:auto
call :run_common Auto -WriteProbe -OpenReport

:fix
call :run_common Fix -WriteProbe -OpenReport -SetMachineHwPath

:compare
call :run_common Compare -OpenReport
goto menu

:editini
start "" notepad.exe "%INI%"
goto menu

:openout
if not exist "%OUTDIR%" mkdir "%OUTDIR%" >nul 2>&1
start "" explorer.exe "%OUTDIR%"
goto menu

:rollback
echo.
set "RBFILE="
set /p RBFILE=Informe o caminho completo do arquivo .reg de backup: 
if not defined RBFILE goto menu
call :run_common Rollback -RollbackFile "%RBFILE%"