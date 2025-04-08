@echo off
echo ===== Criando instalador para DUA Automation =====

REM Compilar o aplicativo primeiro
call python build.py
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha ao compilar o aplicativo.
    goto :fim
)

REM Verificar se o Inno Setup está instalado
reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1" /v DisplayIcon >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Inno Setup não encontrado. Por favor, instale o Inno Setup primeiro.
    echo Você pode baixá-lo em: https://jrsoftware.org/isinfo.php
    goto :fim
)

REM Executar o Inno Setup Compiler
echo Criando instalador...
for /f "tokens=*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Inno Setup 6_is1" /v "InstallLocation"') do (
    for /f "tokens=2*" %%b in ("%%a") do set INNOSETUP=%%c
)
"%INNOSETUP%\ISCC.exe" installer.iss

if %ERRORLEVEL% NEQ 0 (
    echo [ERRO] Falha ao criar o instalador.
) else (
    echo.
    echo =============================================
    echo Instalador criado com sucesso!
    echo Disponível em: %CD%\Output\DUA_Automation_Setup.exe
    echo =============================================
)

:fim
pause
