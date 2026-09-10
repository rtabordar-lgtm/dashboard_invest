@echo off
title Dashboard Taborda - Investimentos
color 0B

echo ======================================================================
echo           INICIANDO DASHBOARD INTERATIVO - TABORDA
echo ======================================================================
echo.

cd /d "%~dp0"

:: 1. Encerrar qualquer processo orfao na porta 8501
for /f "tokens=5" %%a in ('netstat -aon 2^>nul ^| findstr ":8501" ^| findstr "LISTENING"') do (
    echo Liberando porta 8501 - encerrando sessao anterior PID %%a...
    taskkill /F /PID %%a >nul 2>&1
)

:: 2. Verificar se o Python esta instalado
set PYTHON_CMD=python
python --version >nul 2>&1
if %ERRORLEVEL% EQU 0 goto CHECK_VENV

py --version >nul 2>&1
if %ERRORLEVEL% EQU 0 (
    set PYTHON_CMD=py
    goto CHECK_VENV
)

echo [AVISO] O Python nao foi encontrado neste computador!
echo Para rodar o Dashboard Web, instale o Python em: https://www.python.org/downloads/
echo Lembre-se de marcar a opcao "Add Python to PATH" durante a instalacao.
echo.
echo Nota: A planilha Excel funciona normalmente sem o Python.
echo.
pause
exit /b 1

:CHECK_VENV
:: 3. Verificar se o ambiente virtual existe e funciona
if not exist ".venv\Scripts\python.exe" goto BUILD_VENV
".venv\Scripts\python.exe" -c "import sys" >nul 2>&1
if %ERRORLEVEL% NEQ 0 goto BUILD_VENV
goto RUN_APP

:BUILD_VENV
echo [CONFIGURACAO AUTOMATICA] Configurando o ambiente para este computador...
echo Isso acontece apenas na primeira execucao ou ao migrar de PC. Aguarde...
echo.
if exist ".venv" rmdir /s /q ".venv"
echo Criando ambiente virtual...
%PYTHON_CMD% -m venv .venv
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao criar o ambiente virtual.
    pause
    exit /b 1
)
echo Instalando pacotes necessarios...
".venv\Scripts\python.exe" -m pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo Erro ao instalar os pacotes.
    pause
    exit /b 1
)
echo Ambiente configurado com sucesso!
echo.

:RUN_APP
:: 4. Configurar credenciais para nao pedir email
if not exist "%USERPROFILE%\.streamlit" mkdir "%USERPROFILE%\.streamlit" 2>nul
if not exist "%USERPROFILE%\.streamlit\credentials.toml" (
    echo [general] > "%USERPROFILE%\.streamlit\credentials.toml"
    echo email = "" >> "%USERPROFILE%\.streamlit\credentials.toml"
)

:: 5. Iniciar o Dashboard
echo Abrindo painel executivo no seu navegador padrao...
echo Pressione Ctrl+C nesta janela caso queira encerrar o painel.
echo.
".venv\Scripts\python.exe" -m streamlit run app_dashboard.py --browser.gatherUsageStats false

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo Ocorreu um encerramento inesperado.
    pause
)
