@echo off
title AutoSky Inventario v3.70
color 0B
cd /d "%~dp0"

echo.
echo  ================================================
echo   AUTOSKY SISTEMA DE INVENTARIO v3.70
echo   Build: 07/04/2026 GMT-5  ^|  Puerto: 8502
echo  ================================================
echo.
echo  ARCHIVO EJECUTANDO: %~dp0app_web.py
echo  DIRECTORIO: %~dp0
echo.

:: Matar Streamlit en puerto 8501 Y 8502
echo  [1/5] Cerrando versiones anteriores...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8501 " ^| findstr LISTENING') do (
    taskkill /F /PID %%a >nul 2>&1
)
for /f "tokens=5" %%a in ('netstat -aon ^| findstr ":8502 " ^| findstr LISTENING') do (
    taskkill /F /PID %%a >nul 2>&1
)
taskkill /F /IM "streamlit.exe" >nul 2>&1
timeout /t 2 /nobreak >nul

:: Buscar Python
echo  [2/5] Buscando Python...
set PYTHON_CMD=
python --version >nul 2>&1
if not errorlevel 1 set PYTHON_CMD=python
if "%PYTHON_CMD%"=="" (py --version >nul 2>&1 && set PYTHON_CMD=py)
for %%V in (314 313 312 311 310) do (
    if "%PYTHON_CMD%"=="" if exist "%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe" (
        set PYTHON_CMD="%LOCALAPPDATA%\Programs\Python\Python%%V\python.exe"
    )
)
if "%PYTHON_CMD%"=="" (
    echo  [ERROR] Python no encontrado.
    pause & exit /b 1
)
echo  Python: %PYTHON_CMD%

:: Instalar dependencias
echo  [3/5] Verificando dependencias...
%PYTHON_CMD% -m pip install streamlit pandas openpyxl matplotlib reportlab xlrd==2.0.1 --quiet --disable-pip-version-check

:: Limpiar cache de Streamlit
echo  [4/5] Limpiando cache...
rmdir /s /q "%USERPROFILE%\.streamlit\cache" >nul 2>&1
rmdir /s /q "%TEMP%\streamlit" >nul 2>&1

:: Generar log de inicio
set LOGFILE=%~dp0inicio.log
echo ================================================ > "%LOGFILE%"
echo  AutoSky Inventario v3.70 >> "%LOGFILE%"
echo  Inicio: %DATE% %TIME% >> "%LOGFILE%"
echo  Directorio: %~dp0 >> "%LOGFILE%"
echo  Python: %PYTHON_CMD% >> "%LOGFILE%"
echo  Puerto: 8502 >> "%LOGFILE%"
echo  app_web.py: %~dp0app_web.py >> "%LOGFILE%"
echo ================================================ >> "%LOGFILE%"

:: Verificar que app_web.py existe y tiene la version correcta
echo  [5/5] Verificando archivo...
if not exist "%~dp0app_web.py" (
    echo  [ERROR] app_web.py no encontrado en %~dp0 >> "%LOGFILE%"
    echo  [ERROR] app_web.py no encontrado.
    pause & exit /b 1
)
%PYTHON_CMD% -c "import re; c=open('app_web.py',encoding='utf-8').read(); v=re.search('APP_VERSION = \"(.+?)\"',c); print(' Version en app_web.py:', v.group(1) if v else 'NO ENCONTRADA')"
%PYTHON_CMD% -c "import re; c=open('app_web.py',encoding='utf-8').read(); v=re.search('APP_VERSION = \"(.+?)\"',c); open('inicio.log','a',encoding='utf-8').write('Version en app_web.py: '+(v.group(1) if v else 'NO ENCONTRADA')+'\n')" >nul 2>&1

echo.
echo  Abriendo: http://localhost:8502
echo  Para detener: cierre esta ventana
echo  Log guardado en: %~dp0inicio.log
echo.

:: Iniciar Streamlit
%PYTHON_CMD% main_web.py

if errorlevel 1 (
    echo. >> "%LOGFILE%"
    echo  ERROR al iniciar Streamlit >> "%LOGFILE%"
    echo.
    echo  [ERROR] Fallo al iniciar. Ver inicio.log para detalles.
    pause
)
