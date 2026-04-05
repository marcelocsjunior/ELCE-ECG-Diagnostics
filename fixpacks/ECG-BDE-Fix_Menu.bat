@echo off
chcp 65001 >nul
setlocal EnableExtensions EnableDelayedExpansion

set "SCRIPT_DIR=%~dp0"
set "CORE_PS1=%SCRIPT_DIR%BDE-Fix-Core.ps1"
set "STATE_FILE=%TEMP%\ELCE_ECG_SelectedProfile.ini"
set "PROFILES_FILE=%SCRIPT_DIR%ECG_UnitProfiles.json"
set "RUNBOOK=%SCRIPT_DIR%runbookECG-BDE-Fix.txt"
set "LATEST_REPORT=C:\ECG\Output\Latest\ELCE_ECG_Diagnostics_Report.html"

if not exist "%PROFILES_FILE%" set "PROFILES_FILE=%SCRIPT_DIR%..\src\ECG_UnitProfiles.json"
if not exist "%RUNBOOK%" set "RUNBOOK=%SCRIPT_DIR%..\docs\runbooks\runbookECG-BDE-Fix.txt"

if not exist "%CORE_PS1%" (
    echo.
    echo [ERRO] Core unificado nao encontrado:
    echo %CORE_PS1%
    echo.
    pause
    exit /b 1
)

call :load_state

:MENU
cls
echo ============================================
echo   ELCE ECG - Menu Unificado
echo ============================================
echo.
echo Perfil atual........: !PROFILE!
if /I "!PROFILE!"=="CUSTOM" (
  echo DB custom.........: !CUSTOMDB!
  echo NETDIR custom.....: !CUSTOMNETDIR!
)
echo Escopo HW DB........: !HWSCOPE!
echo Profiles file.......: %PROFILES_FILE%
echo.
echo [1] Executar diagnostico completo ^(gera HTML^)
echo [2] Corrigir NETDIR
echo [3] Corrigir HW_CAMINHO_DB
echo [4] Criar diretorios padrao
echo [5] Aplicar todas as correcoes
echo [6] Selecionar / alterar perfil
echo [7] Abrir runbook
echo [8] Abrir ultimo laudo HTML
echo [9] Exibir perfil atual
echo [10] Alterar escopo HW_CAMINHO_DB
echo [0] Sair
echo.
set /p opt="Escolha uma opcao: "

if "%opt%"=="1" goto RUN_DIAG
if "%opt%"=="2" goto RUN_NETDIR
if "%opt%"=="3" goto RUN_HWDB
if "%opt%"=="4" goto RUN_DIRS
if "%opt%"=="5" goto RUN_ALL
if "%opt%"=="6" goto SELECT_PROFILE
if "%opt%"=="7" goto OPEN_RUNBOOK
if "%opt%"=="8" goto OPEN_LAST_REPORT
if "%opt%"=="9" goto SHOW_PROFILE
if "%opt%"=="10" goto CHANGE_SCOPE
if "%opt%"=="0" exit /b 0
goto MENU

:RUN_DIAG
call :run_core DIAG
goto MENU

:RUN_NETDIR
call :run_core NETDIR
goto MENU

:RUN_HWDB
call :run_core HW_CAMINHO_DB
goto MENU

:RUN_DIRS
call :run_core DIRECTORIES
goto MENU

:RUN_ALL
call :run_core ALL
goto MENU

:run_core
set "TASKMODE=%~1"
call :ensure_custom_ready
if errorlevel 1 goto MENU

echo.
echo [INFO] Executando !TASKMODE! para perfil !PROFILE!...
if /I "!PROFILE!"=="CUSTOM" (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CORE_PS1%" -Profile "!PROFILE!" -CustomDbPath "!CUSTOMDB!" -CustomNetDir "!CUSTOMNETDIR!" -TaskMode "!TASKMODE!" -HwScope "!HWSCOPE!" -ProfilesFile "%PROFILES_FILE!" -OpenReportOnSuccess
) else (
    powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%CORE_PS1%" -Profile "!PROFILE!" -TaskMode "!TASKMODE!" -HwScope "!HWSCOPE!" -ProfilesFile "%PROFILES_FILE!" -OpenReportOnSuccess
)
set "LASTCODE=%ERRORLEVEL%"
echo.
if "!LASTCODE!"=="0" (
    echo [OK] Execucao concluida.
) else (
    echo [ERRO] Execucao terminou com ExitCode=!LASTCODE!
)
echo.
pause
exit /b 0

:SELECT_PROFILE
cls
echo ============================================
echo   Selecionar / alterar perfil
echo ============================================
echo.
echo Perfil atual: !PROFILE!
echo.
echo [1] UN1
echo [2] UN2
echo [3] UN3
echo [4] CUSTOM
echo [0] Voltar
echo.
set /p pf="Escolha um perfil: "

if "%pf%"=="1" (
    set "PROFILE=UN1"
    set "CUSTOMDB="
    set "CUSTOMNETDIR="
    call :save_state
    goto MENU
)
if "%pf%"=="2" (
    set "PROFILE=UN2"
    set "CUSTOMDB="
    set "CUSTOMNETDIR="
    call :save_state
    goto MENU
)
if "%pf%"=="3" (
    set "PROFILE=UN3"
    set "CUSTOMDB="
    set "CUSTOMNETDIR="
    call :save_state
    goto MENU
)
if "%pf%"=="4" goto SET_CUSTOM
if "%pf%"=="0" goto MENU
goto SELECT_PROFILE

:SET_CUSTOM
cls
echo ============================================
echo   Configurar perfil CUSTOM
echo ============================================
echo.
set "PROFILE=CUSTOM"
set /p CUSTOMDB="Informe o caminho CUSTOM do DB (UNC ou local): "
if not defined CUSTOMDB (
    echo.
    echo [ERRO] CUSTOMDB nao pode ficar vazio.
    pause
    goto SELECT_PROFILE
)
set /p CUSTOMNETDIR="Informe o caminho CUSTOM do NETDIR (ENTER para derivar DB\NetDir): "
if not defined CUSTOMNETDIR set "CUSTOMNETDIR=!CUSTOMDB!\NetDir"
call :save_state
goto MENU

:OPEN_RUNBOOK
if exist "%RUNBOOK%" (
    start "" notepad.exe "%RUNBOOK%"
) else (
    echo.
    echo [ERRO] Runbook nao encontrado:
    echo %RUNBOOK%
    echo.
    pause
)
goto MENU

:OPEN_LAST_REPORT
if exist "%LATEST_REPORT%" (
    start "" "%LATEST_REPORT%"
) else (
    echo.
    echo [ERRO] Ultimo laudo HTML nao encontrado:
    echo %LATEST_REPORT%
    echo.
    pause
)
goto MENU

:SHOW_PROFILE
cls
echo ============================================
echo   Perfil atualmente selecionado
echo ============================================
echo.
echo PROFILE=!PROFILE!
echo CUSTOMDB=!CUSTOMDB!
echo CUSTOMNETDIR=!CUSTOMNETDIR!
echo HWSCOPE=!HWSCOPE!
echo STATE_FILE=%STATE_FILE%
echo PROFILES_FILE=%PROFILES_FILE%
echo RUNBOOK=%RUNBOOK%
echo LATEST_REPORT=%LATEST_REPORT%
echo.
pause
goto MENU

:CHANGE_SCOPE
cls
echo ============================================
echo   Alterar escopo HW_CAMINHO_DB
echo ============================================
echo.
echo Escopo atual: !HWSCOPE!
echo.
echo [1] User
echo [2] Machine
echo [3] Process
echo [0] Voltar
echo.
set /p sc="Escolha uma opcao: "
if "%sc%"=="1" set "HWSCOPE=User"
if "%sc%"=="2" set "HWSCOPE=Machine"
if "%sc%"=="3" set "HWSCOPE=Process"
if "%sc%"=="0" goto MENU
call :save_state
goto MENU

:ensure_custom_ready
if /I not "!PROFILE!"=="CUSTOM" exit /b 0
if not defined CUSTOMDB (
    echo.
    echo [ERRO] Perfil CUSTOM sem CUSTOMDB definido.
    pause
    exit /b 1
)
if not defined CUSTOMNETDIR (
    echo.
    echo [ERRO] Perfil CUSTOM sem CUSTOMNETDIR definido.
    pause
    exit /b 1
)
exit /b 0

:load_state
if not exist "%STATE_FILE%" (
    set "PROFILE=UN1"
    set "CUSTOMDB="
    set "CUSTOMNETDIR="
    set "HWSCOPE=User"
    call :save_state
    goto :eof
)

set "PROFILE=UN1"
set "CUSTOMDB="
set "CUSTOMNETDIR="
set "HWSCOPE=User"

for /f "usebackq tokens=1* delims==" %%A in ("%STATE_FILE%") do (
    if /I "%%A"=="PROFILE" set "PROFILE=%%B"
    if /I "%%A"=="CUSTOMDB" set "CUSTOMDB=%%B"
    if /I "%%A"=="CUSTOMNETDIR" set "CUSTOMNETDIR=%%B"
    if /I "%%A"=="HWSCOPE" set "HWSCOPE=%%B"
)
goto :eof

:save_state
> "%STATE_FILE%" echo PROFILE=!PROFILE!
>> "%STATE_FILE%" echo CUSTOMDB=!CUSTOMDB!
>> "%STATE_FILE%" echo CUSTOMNETDIR=!CUSTOMNETDIR!
>> "%STATE_FILE%" echo HWSCOPE=!HWSCOPE!
goto :eof
