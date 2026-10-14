@echo off
chcp 65001 >nul
title Instalar Menu Unificar - Império do Extintor
echo ================================================================
echo   INSTALANDO MENU 'UNIFICAR' NO BOTÃO DIREITO DO WINDOWS...
echo ================================================================
echo.

cd /d "%~dp0"

if exist "unificar.exe" (
    start /wait "" "unificar.exe" --install
) else (
    python unificar.py --install
)

echo.
echo ================================================================
echo   Menu registrado com sucesso com o novo ícone estilizado!
echo ================================================================
echo.
echo No Windows 11, por padrão, opções do botão direito ficam dentro
echo de "Mostrar mais opções".
echo.
set /p WIN11_OPT="Deseja ativar a abertura DIRETA no 1º clique (sem 'Mostrar mais opções')? (S/N): "
if /i "%WIN11_OPT%"=="S" (
    echo.
    echo Ativando menu direto no Windows 11...
    reg add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve >nul 2>&1
    echo Reiniciando Windows Explorer...
    taskkill /f /im explorer.exe >nul 2>&1
    timeout /t 1 >nul
    start explorer.exe
    echo Menu direto ativado com sucesso!
)

echo.
echo ================================================================
echo   Instalação finalizada!
echo ================================================================
echo.
pause
