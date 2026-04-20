@echo off
title Crear Ejecutable - Sistema de Inventario v3.35
color 0B
cd /d "%~dp0"

echo ================================================
echo   CREAR EJECUTABLE WINDOWS (.exe)
echo   Sistema de Inventario v3.35
echo ================================================
echo.
echo Este script compilara el sistema en un .exe
echo que funciona SIN necesitar Python instalado.
echo.
echo Tiempo estimado: 3-5 minutos.
echo.
pause

:: ── Buscar Python ──────────────────────────────────────────────
set PYTHON_CMD=

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
    echo [ERROR] Python no encontrado. Instale Python desde https://python.org
    pause
    exit /b 1
)

echo Python: 
%PYTHON_CMD% --version
echo.

:: ── Instalar PyInstaller y dependencias ────────────────────────
echo [1/4] Instalando PyInstaller...
%PYTHON_CMD% -m pip install pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo [ERROR] No se pudo instalar PyInstaller.
    pause
    exit /b 1
)

echo [2/4] Instalando dependencias del sistema...
%PYTHON_CMD% -m pip install customtkinter pandas openpyxl matplotlib reportlab --quiet --disable-pip-version-check

:: ── Detectar rutas de customtkinter para incluirlas ───────────
echo [3/4] Detectando rutas de dependencias...
for /f "delims=" %%i in ('%PYTHON_CMD% -c "import customtkinter; import os; print(os.path.dirname(customtkinter.__file__))"') do set CTK_PATH=%%i

:: ── Compilar ───────────────────────────────────────────────────
echo [4/4] Compilando ejecutable (puede tardar varios minutos)...
echo.

%PYTHON_CMD% -m PyInstaller ^
    --onefile ^
    --windowed ^
    --name "SistemaInventario_v3.35" ^
    --add-data "app;app" ^
    --add-data "%CTK_PATH%;customtkinter" ^
    --hidden-import "customtkinter" ^
    --hidden-import "pandas" ^
    --hidden-import "openpyxl" ^
    --hidden-import "matplotlib" ^
    --hidden-import "matplotlib.backends.backend_tkagg" ^
    --hidden-import "reportlab" ^
    --hidden-import "PIL" ^
    --hidden-import "PIL._tkinter_finder" ^
    --collect-all "customtkinter" ^
    --collect-all "matplotlib" ^
    main.py

if errorlevel 1 (
    echo.
    echo [ERROR] La compilacion fallo. Revise los mensajes anteriores.
    pause
    exit /b 1
)

echo.
echo ================================================
echo   COMPILACION EXITOSA
echo ================================================
echo.
echo El ejecutable se encuentra en:
echo   dist\SistemaInventario_v3.35.exe
echo.
echo Puede copiar ese .exe a cualquier PC con Windows
echo sin necesitar Python instalado.
echo.
echo NOTA: El primer inicio puede tardar 10-20 segundos
echo       mientras el ejecutable se descomprime.
echo.
explorer dist
pause
