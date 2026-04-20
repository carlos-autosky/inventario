@echo off
title Sistema de Inventario v3.35
color 0A
cd /d "%~dp0"

echo ================================================
echo   SISTEMA DE INVENTARIO v3.35
echo   Iniciando...
echo ================================================
echo.

:: ── Buscar Python ──────────────────────────────────────────────
set PYTHON_CMD=

python --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=python

if "%PYTHON_CMD%"=="" (
    py --version >nul 2>&1
    if not errorlevel 1 set PYTHON_CMD=py
)

:: Rutas comunes de instalación
for %%V in (314 313 312 311 310) do (
    if "%PYTHON_CMD%"=="" (
        if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
            set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
        )
    )
)
for %%V in (314 313 312 311 310) do (
    if "%PYTHON_CMD%"=="" (
        if exist "C:\Python%%V\python.exe" (
            set PYTHON_CMD="C:\Python%%V\python.exe"
        )
    )
)

if "%PYTHON_CMD%"=="" (
    echo [ERROR] Python no encontrado.
    echo.
    echo Instale Python desde: https://python.org/downloads
    echo IMPORTANTE: Marque la opcion "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

echo Python encontrado:
%PYTHON_CMD% --version
echo.

:: ── Instalar dependencias ───────────────────────────────────────
echo Verificando dependencias...
%PYTHON_CMD% -m pip install --upgrade pip --quiet --disable-pip-version-check 2>nul
%PYTHON_CMD% -m pip install customtkinter>=5.2.0 pandas>=2.0.0 openpyxl>=3.1.0 matplotlib>=3.7.0 reportlab>=4.0.0 --quiet --disable-pip-version-check

if errorlevel 1 (
    echo [ERROR] No se pudieron instalar las dependencias.
    echo Verifique su conexion a internet e intente nuevamente.
    pause
    exit /b 1
)

echo Dependencias OK.
echo.

:: ── Iniciar sistema ─────────────────────────────────────────────
echo Iniciando Sistema de Inventario...
%PYTHON_CMD% main.py

if errorlevel 1 (
    echo.
    echo ================================================
    echo   Error al iniciar. Revise el mensaje anterior.
    echo ================================================
    pause
)
