@echo off
chcp 65001 >nul
title Desinstalar Menu Unificar - Império do Extintor
echo ================================================================
echo   REMOVENDO MENU 'UNIFICAR' DO BOTÃO DIREITO DO WINDOWS...
echo ================================================================
echo.

cd /d "%~dp0"

if exist "unificar.exe" (
    start "" "unificar.exe" --uninstall
) else (
    python unificar.py --uninstall
)

echo.
echo ================================================================
echo   Concluído! A opção "Unificar" foi removida do menu de contexto.
echo ================================================================
echo.
pause
