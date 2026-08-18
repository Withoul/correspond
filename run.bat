@echo off
title Correspon - Automatizacion de Correspondencia
chcp 65001 > nul
echo.
echo ========================================================
echo   Iniciando Correspon (Servidor + Interfaz de Usuario)
echo ========================================================
echo.
cd /d "%~dp0"
python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo ejecutar 'python main.py'.
    echo Verifique que Python 3 esta instalado en el sistema.
    pause
)
