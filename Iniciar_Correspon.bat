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

:: 2. VERIFICAR PYTHON EN EL SISTEMA
python --version >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Python no se detecta en el comando 'python'. Intentando con 'py'...
    py --version >nul 2>&1
    if not errorlevel 1 (
        doskey python=py $*
    ) else (
        echo [AVISO] Python 3 no se encuentra instalado en este equipo.
        echo Intentando instalación automática via winget...
        echo.
        winget install --id Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
        if errorlevel 1 (
            echo.
            echo [ERROR] No se pudo instalar Python 3 automáticamente.
            echo Por favor descarga e instala Python 3 desde https://www.python.org/ (asegúrate de marcar 'Add Python to PATH') y vuelve a ejecutar este archivo.
            pause
            exit /b 1
        )
        set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"
    )
)

echo [OK] Python detectado correctamente.

:: 3. VERIFICAR E INSTALAR DEPENDENCIAS SI FALTAN
python -c "import fastapi, uvicorn, openpyxl, pandas, docxtpl, docx, mammoth, win32com" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Detectadas dependencias faltantes. Instalando paquetes requeridos...
    if exist "wheels" (
        python -m pip install -r requirements.txt --find-links="wheels" --no-index >nul 2>&1
    )
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Ocurrió un problema al instalar las dependencias con pip.
        pause
        exit /b 1
    )
    echo [OK] Dependencias instaladas exitosamente.
) else (
    echo [OK] Dependencias verificadas y listas.
)

:: 4. CREAR O ACTUALIZAR ACCESO DIRECTO EN EL ESCRITORIO
if exist "app_icon.ico" (
    powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Correspon.lnk')); $s.TargetPath = '%~dp0Iniciar_Correspon.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = '%~dp0app_icon.ico'; $s.Save()" >nul 2>&1
)

echo.
echo ========================================================
echo   INICIANDO SERVIDOR Y APLICACIÓN...
echo ========================================================
echo.

python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Ocurrió un error al iniciar Correspon.
    pause
)
