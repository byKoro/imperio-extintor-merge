@echo off
chcp 65001 >nul
title Restaurar Menu Padrão do Windows 11 - Império do Extintor
echo ================================================================
echo   RESTAURANDO MENU PADRÃO DO WINDOWS 11
echo ================================================================
echo.
echo O Windows Explorer será reiniciado para restaurar o menu original.
echo.
pause

reg delete "HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}" /f >nul 2>&1

echo.
echo Reiniciando o Windows Explorer...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 1 >nul
start explorer.exe

echo.
echo ================================================================
echo   Pronto! Menu padrão do Windows 11 restaurado.
echo ================================================================
echo.
pause
