@echo off
title Instalador - Sistema de Inventario v3.35
color 0A
cd /d "%~dp0"

echo ================================================
echo   INSTALADOR - SISTEMA DE INVENTARIO v3.35
echo ================================================
echo.
echo Este instalador:
echo   1. Verifica/instala Python si es necesario
echo   2. Instala las dependencias del sistema
echo   3. Crea un acceso directo en el Escritorio
echo   4. Inicia el sistema
echo.
pause

:: ── Verificar Python ───────────────────────────────────────────
set PYTHON_CMD=
set NEEDS_PYTHON=0

python --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=python

if "%PYTHON_CMD%"=="" (
    py --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=py
)

for %%V in (314 313 312 311 310) do (
    if "%PYTHON_CMD%"=="" (
        if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
            set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        )
    )
)

if "%PYTHON_CMD%"=="" (
    echo [!] Python no esta instalado.
    echo.
    echo Abriendo la pagina de descarga de Python...
    start https://www.python.org/ftp/python/3.13.0/python-3.13.0-amd64.exe
    echo.
    echo IMPORTANTE durante la instalacion:
    echo   - Marque "Add Python to PATH"
    echo   - Haga clic en "Install Now"
    echo.
    echo Una vez instalado Python, ejecute este instalador nuevamente.
    pause
    exit /b 1
)

echo [OK] Python encontrado:
%PYTHON_CMD% --version
echo.

:: ── Instalar dependencias ───────────────────────────────────────
echo Instalando dependencias (requiere internet)...
%PYTHON_CMD% -m pip install --upgrade pip --quiet --disable-pip-version-check 2>nul
%PYTHON_CMD% -m pip install customtkinter pandas openpyxl matplotlib reportlab --quiet --disable-pip-version-check

if errorlevel 1 (
    echo [ERROR] Fallo la instalacion de dependencias.
    echo Verifique su conexion a internet.
    pause
    exit /b 1
)
echo [OK] Dependencias instaladas.
echo.

:: ── Crear acceso directo en el Escritorio ──────────────────────
echo Creando acceso directo en el Escritorio...
set SCRIPT_DIR=%~dp0
set DESKTOP=%USERPROFILE%\Desktop

:: Usar PowerShell para crear el acceso directo
powershell -Command ^
    "$ws = New-Object -ComObject WScript.Shell; ^
     $sc = $ws.CreateShortcut('%DESKTOP%\Sistema Inventario v3.35.lnk'); ^
     $sc.TargetPath = '%SCRIPT_DIR%run.bat'; ^
     $sc.WorkingDirectory = '%SCRIPT_DIR%'; ^
     $sc.Description = 'Sistema de Inventario v3.35'; ^
     $sc.Save()" 2>nul

if exist "%DESKTOP%\Sistema Inventario v3.35.lnk" (
    echo [OK] Acceso directo creado en el Escritorio.
) else (
    echo [!] No se pudo crear el acceso directo. Use run.bat directamente.
)
echo.

:: ── Iniciar sistema ─────────────────────────────────────────────
echo ================================================
echo   Instalacion completada. Iniciando sistema...
echo ================================================
echo.
timeout /t 2 /nobreak >nul
%PYTHON_CMD% main.py

if errorlevel 1 (
    echo.
    echo Error al iniciar. Revise el mensaje anterior.
    pause
)
