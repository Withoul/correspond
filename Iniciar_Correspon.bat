@echo off
title Correspon - Sistema de Correspondencia Masiva
chcp 65001 > nul
cls

echo ========================================================
echo   CORRESPON - SISTEMA DE CORRESPONDENCIA MASIVA
echo ========================================================
echo.

cd /d "%~dp0"

:: 1. VERIFICAR SI EXISTE EJECUTABLE COMPILADO
if exist "dist\Correspon\Correspon.exe" (
    echo [OK] Ejecutable detectado. Iniciando Correspon.exe...
    start "" "dist\Correspon\Correspon.exe"
    exit /b 0
)

:: 2. DETECTAR EL INTERPRETE DE PYTHON ADECUADO
set "PY_CMD="

if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :python_found
)

if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    goto :python_found
)

if exist "%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\hermes\hermes-agent\venv\Scripts\python.exe"
    goto :python_found
)

py -3.11 --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py -3.11"
    goto :python_found
)

py --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=py"
    goto :python_found
)

python --version >nul 2>&1
if not errorlevel 1 (
    set "PY_CMD=python"
    goto :python_found
)

:install_python
echo [AVISO] Python 3 no se encuentra instalado en este equipo.
echo Intentando instalacion automatica via winget...
winget install --id Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" (
    set "PY_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    goto :python_found
)
echo [ERROR] No se pudo detectar Python 3.
echo Descargue e instale Python 3.11 desde python.org y vuelva a intentar.
pause
exit /b 1

:python_found
echo [OK] Interprete detectado correctamente.

:: 3. VERIFICAR E INSTALAR DEPENDENCIAS SI FALTAN
"%PY_CMD%" -c "import fastapi, uvicorn, openpyxl, pandas, docxtpl, docx, mammoth, win32com, webview" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Detectadas dependencias faltantes. Instalando paquetes requeridos...
    if exist "wheels" (
        echo [INFO] Instalando desde paquete local wheels...
        "%PY_CMD%" -m pip install -r requirements.txt --find-links="wheels" --no-index
    ) else (
        "%PY_CMD%" -m pip install -r requirements.txt
    )
    if errorlevel 1 (
        echo [ERROR] Ocurrio un problema al instalar las dependencias con pip.
        pause
        exit /b 1
    )
    echo [OK] Dependencias instaladas exitosamente.
) else (
    echo [OK] Todas las dependencias estan verificadas y listas.
)

:: 4. CREAR O ACTUALIZAR ACCESO DIRECTO EN EL ESCRITORIO
if exist "app_icon.ico" (
    powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Correspon.lnk')); $s.TargetPath = '%~dp0Iniciar_Correspon.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = '%~dp0app_icon.ico'; $s.Save()" >nul 2>&1
)

echo.
echo ========================================================
echo   INICIANDO CORRESPON...
echo ========================================================
echo.

"%PY_CMD%" main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Ocurrio un error al ejecutar la aplicacion.
    pause
)
