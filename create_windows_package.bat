@echo off
:: filepath: /home/gabriel/Documents/DUA_automation/dua/create_windows_package.bat
:: Script para compilar o DUA Automation e criar instalador para Windows

echo ===== DUA Automation - Script de Build para Windows =====
echo.

:: Verificar se Python está instalado
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Python nao encontrado. Por favor, instale Python 3.6 ou superior.
    goto :fim
)

:: Verificar versão do Python
for /f "tokens=2" %%i in ('python --version 2^>^&1') do set PYTHON_VERSION=%%i
echo Python versao %PYTHON_VERSION% encontrado

:: Verificar ambiente virtual
if exist venv\ (
    echo Ambiente virtual encontrado
    call venv\Scripts\activate
) else (
    echo Criando ambiente virtual...
    python -m venv venv
    call venv\Scripts\activate
    echo Ambiente virtual ativado
)

:: Instalar dependências
echo.
echo Instalando dependencias...
pip install -r requirements.txt
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha ao instalar dependencias.
    goto :fim
)

:: Executar o script de build
echo.
echo Iniciando build do executavel...
python build.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha ao executar o build.
    goto :fim
)

:: Verificar se o instalador foi gerado
if exist Output\*.exe (
    echo.
    echo =============================================
    echo Build concluido com sucesso!
    echo Instalador disponivel em: %CD%\Output\
    echo =============================================
) else (
    if exist dist\DUA_Automation.exe (
        echo.
        echo =============================================
        echo Build concluido com sucesso!
        echo Executavel disponivel em: %CD%\dist\DUA_Automation.exe
        echo =============================================
    )
)

:fim
pause