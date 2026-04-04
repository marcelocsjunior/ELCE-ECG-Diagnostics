@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%ELCE_ECG_Diagnostics.ps1"

if not exist "%SCRIPT%" if exist "C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1" set "SCRIPT=C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo [ERRO] Script principal nao encontrado:
    echo %SCRIPT%
    echo.
    pause
    exit /b 1
)

echo Executando diagnostico ECG...
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -StagePriority ABRIR_EXAME -SymptomCode LENTIDAO_TRAVAMENTO -ObservationMinutes 10 -SampleIntervalSeconds 20 -OpenReportOnSuccess
echo.
echo Diagnostico concluido. Pressione qualquer tecla para sair...
pause > nul
