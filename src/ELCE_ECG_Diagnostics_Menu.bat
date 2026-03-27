@echo off
chcp 65001 >nul
setlocal EnableExtensions

set "TOOL_NAME=ELCE ECG Diagnostics"
set "TOOL_DIR=C:\ECG\Tool"
set "SCRIPT=%TOOL_DIR%\ELCE_ECG_Diagnostics.ps1"
set "OUTPUT_DIR=C:\ECG\Output"
set "LATEST_REPORT=%OUTPUT_DIR%\Latest\ELCE_ECG_Diagnostics_Report.html"

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
echo [0] Sair
echo.
set /p opt="Escolha uma opcao: "

if "%opt%"=="1" goto RUN
if "%opt%"=="2" goto OPEN_REPORT
if "%opt%"=="3" goto OPEN_OUTPUT
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
