@echo off
title Correspon - Ejecutor
chcp 65001 > nul
cd /d "%~dp0"
echo ========================================================
echo   Iniciando Correspon (Servidor Local IP/Puerto)...
echo ========================================================
echo.
python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] No se pudo iniciar el servidor de Correspon.
    echo Si es la primera vez que usas la app en esta PC, ejecuta 'instalar.bat' primero.
    pause
)
