@echo off
setlocal EnableExtensions DisableDelayedExpansion
title ECG Diagnostics AI Prototype v0.2.1

set "SCRIPT_DIR=%~dp0"
set "AI_PS1=%SCRIPT_DIR%ECG_Diagnostics_AI_Prototype_v0_2_1.ps1"
set "PROFILE_INI=%SCRIPT_DIR%ECG_FieldKit_Unified_v6_3_2.ini"
set "CONFIG_JSON=%SCRIPT_DIR%ECG_AI.config.json"
set "PROMPTS_JSON=%SCRIPT_DIR%ECG_AI_Prompts.json"
set "PS_EXE=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"

if not exist "%PS_EXE%" (
  echo [ERRO] powershell.exe nao encontrado.
  pause
  exit /b 1
)
if not exist "%AI_PS1%" (
  echo [ERRO] Arquivo ausente:
  echo %AI_PS1%
  pause
  exit /b 1
)

:MENU
cls
echo ===============================================================
echo   ECG Diagnostics AI Prototype v0.2.1 - off-repo / local rules
echo ===============================================================
echo.
echo 1^) Explicar ultimo laudo
echo 2^) Gerar resumo executivo do ultimo laudo
echo 3^) Gerar parecer tecnico do ultimo laudo
echo 4^) Comparar dois laudos com narrativa
echo 5^) Perguntar sobre o ultimo laudo
echo 6^) Abrir pasta do pacote
echo 7^) Abrir config IA
echo 8^) Abrir catalogo de prompts
echo.
echo 0^) Sair
echo.
set /p OPT=Escolha uma opcao [0-8]: 

if "%OPT%"=="1" goto EXPLAIN
if "%OPT%"=="2" goto EXECUTIVE
if "%OPT%"=="3" goto TECHNICAL
if "%OPT%"=="4" goto COMPARE
if "%OPT%"=="5" goto ASK
if "%OPT%"=="6" goto OPENFOLDER
if "%OPT%"=="7" goto OPENCONFIG
if "%OPT%"=="8" goto OPENPROMPTS
if "%OPT%"=="0" exit /b 0

echo.
echo [ERRO] Opcao invalida.
pause
goto MENU

:RUNAI
echo.
"%PS_EXE%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%AI_PS1%" %RUN_ARGS%
set "RC=%ERRORLEVEL%"
echo.
if "%RC%"=="0" (
  echo [OK] Analise concluida.
) else (
  echo [ERRO] Analise retornou codigo %RC%.
)
echo.
pause
goto MENU

:EXPLAIN
set "RUN_ARGS=-Mode Explain -ProfilePath ""%PROFILE_INI%"" -ConfigPath ""%CONFIG_JSON%"" -PromptCatalogPath ""%PROMPTS_JSON%"" -OpenOutput"
goto RUNAI

:EXECUTIVE
set "RUN_ARGS=-Mode Executive -ProfilePath ""%PROFILE_INI%"" -ConfigPath ""%CONFIG_JSON%"" -PromptCatalogPath ""%PROMPTS_JSON%"" -OpenOutput"
goto RUNAI

:TECHNICAL
echo.
echo Publico do parecer:
echo 1^) DEV
echo 2^) INFRA
echo 3^) CAMPO
echo 4^) EXEC
echo.
set /p AUD=Escolha [1-4]: 
set "AUDIENCE=INFRA"
if "%AUD%"=="1" set "AUDIENCE=DEV"
if "%AUD%"=="2" set "AUDIENCE=INFRA"
if "%AUD%"=="3" set "AUDIENCE=CAMPO"
if "%AUD%"=="4" set "AUDIENCE=EXEC"
set "RUN_ARGS=-Mode Technical -Audience %AUDIENCE% -ProfilePath ""%PROFILE_INI%"" -ConfigPath ""%CONFIG_JSON%"" -PromptCatalogPath ""%PROMPTS_JSON%"" -OpenOutput"
goto RUNAI

:COMPARE
echo.
echo Informe os caminhos completos de dois arquivos JSON OU duas pastas de rodada.
echo.
set "COMPARE_LEFT="
set "COMPARE_RIGHT="
set /p COMPARE_LEFT=Relatorio esquerdo (.json): 
set /p COMPARE_RIGHT=Relatorio direito (.json): 
if "%COMPARE_LEFT%"=="" goto COMPARE_ERROR
if "%COMPARE_RIGHT%"=="" goto COMPARE_ERROR
set "RUN_ARGS=-Mode CompareAI -LeftReportPath ""%COMPARE_LEFT%"" -RightReportPath ""%COMPARE_RIGHT%"" -ProfilePath ""%PROFILE_INI%"" -ConfigPath ""%CONFIG_JSON%"" -PromptCatalogPath ""%PROMPTS_JSON%"" -OpenOutput"
goto RUNAI

:COMPARE_ERROR
echo.
echo [ERRO] Informe os dois caminhos dos relatorios JSON ou das pastas de rodada.
pause
goto MENU

:ASK
echo.
set /p QUESTION=Pergunta sobre o ultimo laudo: 
if "%QUESTION%"=="" (
  echo [ERRO] Pergunta nao informada.
  pause
  goto MENU
)
set "RUN_ARGS=-Mode Ask -Question ""%QUESTION%"" -ProfilePath ""%PROFILE_INI%"" -ConfigPath ""%CONFIG_JSON%"" -PromptCatalogPath ""%PROMPTS_JSON%"" -OpenOutput"
goto RUNAI

:OPENFOLDER
start "" "%SCRIPT_DIR%"
goto MENU

:OPENCONFIG
start "" notepad.exe "%CONFIG_JSON%"
goto MENU

:OPENPROMPTS
start "" notepad.exe "%PROMPTS_JSON%"
goto MENU
