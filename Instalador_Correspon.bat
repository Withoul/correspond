@echo off
title Correspon - Instalador y Ejecutor Principal
chcp 65001 > nul
cls

echo ========================================================
echo   CORRESPON - SISTEMA DE CORRESPONDENCIA MASIVA
echo ========================================================
echo.

cd /d "%~dp0"

:: 1. VERIFICAR SI EXE COMPILADO EXISTE
if exist "dist\Correspon\Correspon.exe" (
    echo [OK] Ejecutable detectado. Iniciando Correspon.exe...
    start "" "dist\Correspon\Correspon.exe"
    exit /b 0
)

:: 2. VERIFICAR SI PYTHON ESTA INSTALADO
python --version >nul 2>&1
if errorlevel 1 (
    echo [AVISO] Python 3 no se encuentra instalado en este equipo.
    echo Intentando instalacion automatica de Python via winget...
    echo.
    winget install --id Python.Python.3.11 --silent --accept-package-agreements --accept-source-agreements
    if errorlevel 1 (
        echo.
        echo [ERROR] No se pudo instalar Python 3 automaticamente.
        echo Por favor descarga e instala Python 3 desde https://www.python.org/ y vuelve a ejecutar.
        pause
        exit /b 1
    )
    set "PATH=%LOCALAPPDATA%\Programs\Python\Python311;%LOCALAPPDATA%\Programs\Python\Python311\Scripts;%PATH%"
)

echo [OK] Python detectado correctamente.

:: 3. VERIFICAR SI LAS LIBRERIAS YA ESTAN INSTALADAS
python -c "import fastapi, uvicorn, openpyxl, pandas, docxtpl, docx, mammoth, win32com" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Instalando librerias requeridas...
    if exist "wheels" (
        python -m pip install -r requirements.txt --find-links="wheels" --no-index >nul 2>&1
    )
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Hubo un problema al instalar las librerias.
        pause
        exit /b 1
    )
    echo [OK] Librerias instaladas exitosamente.
) else (
    echo [OK] Librerias necesarias ya instaladas. Ejecutando de inmediato...
)

:: 4. CREAR ACCESO DIRECTO EN EL ESCRITORIO CON EL ICONO DE LA APP
if exist "app_icon.ico" (
    powershell -Command "$ws = New-Object -ComObject WScript.Shell; $s = $ws.CreateShortcut([System.IO.Path]::Combine([Environment]::GetFolderPath('Desktop'), 'Correspon.lnk')); $s.TargetPath = '%~dp0Instalador_Correspon.bat'; $s.WorkingDirectory = '%~dp0'; $s.IconLocation = '%~dp0app_icon.ico'; $s.Save()" >nul 2>&1
)

echo.
echo ========================================================
echo   INICIANDO APLICACION...
echo ========================================================
echo.

python main.py
if errorlevel 1 (
    echo.
    echo [ERROR] Ocurrio un error al ejecutar la aplicacion.
    pause
)
