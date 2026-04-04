@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "TOOL_NAME=ELCE ECG Diagnostics"
set "SCRIPT_DIR=%~dp0"
set "SCRIPT=%SCRIPT_DIR%ELCE_ECG_Diagnostics.ps1"
set "OUTPUT_DIR=C:\ECG\Output"
set "LATEST_REPORT=%OUTPUT_DIR%\Latest\ELCE_ECG_Diagnostics_Report.html"
set "FIX_MENU=%SCRIPT_DIR%ECG-BDE-Fix_Menu.bat"
set "RUNBOOK=%SCRIPT_DIR%runbookECG-BDE-Fix.txt"

if not exist "%SCRIPT%" if exist "C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1" set "SCRIPT=C:\ECG\Tool\ELCE_ECG_Diagnostics.ps1"
if not exist "%FIX_MENU%" set "FIX_MENU=%SCRIPT_DIR%..\fixpacks\ECG-BDE-Fix_Menu.bat"
if not exist "%RUNBOOK%" set "RUNBOOK=%SCRIPT_DIR%..\docs\runbooks\runbookECG-BDE-Fix.txt"

if not exist "%SCRIPT%" (
    echo.
    echo [ERRO] Script principal nao encontrado:
    echo %SCRIPT%
    echo.
    pause
    exit /b 1
)

:MENU
cls
echo ============================================
echo   %TOOL_NAME%
echo ============================================
echo.
echo [1] Executar diagnostico padrao
echo [2] Abrir ultimo laudo
echo [3] Abrir pasta de saida
echo [4] Abrir menu ECG BDE Fix
echo [5] Abrir runbook BDE
echo [0] Sair
echo.
set /p opt="Escolha uma opcao: "

if "%opt%"=="1" goto RUN
if "%opt%"=="2" goto OPEN_REPORT
if "%opt%"=="3" goto OPEN_OUTPUT
if "%opt%"=="4" goto OPEN_FIX_MENU
if "%opt%"=="5" goto OPEN_RUNBOOK
if "%opt%"=="0" exit /b 0
goto MENU

:RUN
echo.
echo Iniciando diagnostico padrao:
echo   - Etapa priorizada: Abrir exame
echo   - Sintoma: Lentidao / travamentos
echo   - Janela observacional: 10 minutos
echo   - Intervalo entre amostras: 20 segundos
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -StagePriority ABRIR_EXAME -SymptomCode LENTIDAO_TRAVAMENTO -ObservationMinutes 10 -SampleIntervalSeconds 20 -OpenReportOnSuccess
set "EXITCODE=%ERRORLEVEL%"

echo.
if "%EXITCODE%"=="0" (
    echo [OK] Rodada concluida.
) else (
    echo [ERRO] A rodada terminou com falha. ExitCode=%EXITCODE%
)
echo.
pause
goto MENU

:OPEN_REPORT
if exist "%LATEST_REPORT%" (
    start "" "%LATEST_REPORT%"
) else (
    echo.
    echo Nenhum laudo encontrado em:
    echo %LATEST_REPORT%
    echo.
    pause
)
goto MENU

:OPEN_OUTPUT
if exist "%OUTPUT_DIR%" (
    start "" "%OUTPUT_DIR%"
) else (
    echo.
    echo Pasta de saida nao encontrada:
    echo %OUTPUT_DIR%
    echo.
    pause
)
goto MENU

:OPEN_FIX_MENU
if exist "%FIX_MENU%" (
    call "%FIX_MENU%"
) else (
    echo.
    echo Menu de correcao nao encontrado:
    echo %FIX_MENU%
    echo.
    pause
)
goto MENU

:OPEN_RUNBOOK
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
