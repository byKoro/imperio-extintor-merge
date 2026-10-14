@echo off
chcp 65001 >nul
title Ativar Menu Direto no Windows 11 - Império do Extintor
echo ================================================================
echo   ATIVANDO MENU DE CONTEXTO DIRETO NO WINDOWS 11 (1º CLIQUE)
echo ================================================================
echo.
echo Este comando faz com que o menu do botão direito abra diretamente
echo com todas as opções (incluindo "Unificar"), sem precisar clicar
echo em "Mostrar mais opções".
echo.
echo O Windows Explorer será reiniciado para aplicar a alteração.
echo.
pause

reg add "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32" /f /ve

echo.
echo Reiniciando o Windows Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 >nul
start explorer.exe

echo.
echo ================================================================
echo   Pronto! Menu direto ativado com sucesso!
echo   Ao clicar com botão direito, o menu completo abrirá no 1º clique.
echo ================================================================
echo.
pause
