@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM ============================================================
REM Md2PPT.bat - Lanzador automatico de Md2PPT.py
REM Lee la configuracion de Settings.json
REM Busca el primer .md y .pptx en docs\
REM Salida: titulo H1 del markdown o "resultado.pptx"
REM ============================================================

set "SCRIPT_DIR=%~dp0"
set "SETTINGS=%SCRIPT_DIR%Settings.json"
set "PYTHON_EXE="

REM --- Leer InterpreterPath del JSON ---
if not exist "%SETTINGS%" (
    echo [ERROR] No se encontro Settings.json
    echo Usando "python" por defecto...
    set "PYTHON_EXE=python"
) else (
    for /f "usebackq delims=" %%A in (`findstr /i "InterpreterPath" "%SETTINGS%"`) do (
        set "RAW_LINE=%%A"
    )
    REM Extraer solo la ruta usando PowerShell inline
    for /f "usebackq delims=" %%P in (`powershell -NoProfile -Command "(Get-Content '%SETTINGS%' | ConvertFrom-Json).Python.InterpreterPath"`) do (
        set "PYTHON_EXE=%%P"
    )
)

REM --- Leer CompanyName del JSON (puede estar vacio) ---
set "COMPANY_NAME="
for /f "usebackq delims=" %%C in (`powershell -NoProfile -Command "(Get-Content '%SETTINGS%' | ConvertFrom-Json).General.CompanyName"`) do (
    set "COMPANY_NAME=%%C"
)

REM --- Verificar que el intérprete existe ---
if not exist "!PYTHON_EXE!" (
    echo [ERROR] No se encontró el intérprete: !PYTHON_EXE!
    pause
    exit /b 1
)

set "DOCS_DIR=%SCRIPT_DIR%docs"

REM --- Obtener archivos y nombre de salida via PowerShell ---
for /f "usebackq tokens=1,2,3 delims=|" %%A in (`powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%GetTitle.ps1" "%DOCS_DIR%"`) do (
    set "MD_FILE=%%A"
    set "PLANTILLA_FILE=%%B"
    set "SALIDA=%%C"
)

REM --- Verificar que se obtuvieron los valores ---
if "!MD_FILE!"=="" (
    echo [ERROR] No se encontro ningun archivo .md en docs\
    pause
    exit /b 1
)
if "!PLANTILLA_FILE!"=="" (
    echo [ERROR] No se encontro ningun archivo .pptx en docs\
    pause
    exit /b 1
)

echo ============================================================
echo  Md2PPT - Conversor Markdown a PowerPoint
echo ============================================================
echo  Interprete: !PYTHON_EXE!
echo  Entrada:    docs\!MD_FILE!
echo  Plantilla:  docs\!PLANTILLA_FILE!
echo  Salida:     docs\!SALIDA!.pptx
echo ============================================================
echo.

REM --- Limpiar la plantilla (genera una version temporal con solo los 4 layouts) ---
set "PLANTILLA_TEMP=%DOCS_DIR%\_plantilla_tmp.pptx"
echo  [1/2] Limpiando plantilla...
"!PYTHON_EXE!" "%SCRIPT_DIR%LimpiarPlantilla.py" "%DOCS_DIR%\!PLANTILLA_FILE!" "!PLANTILLA_TEMP!"
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Fallo al limpiar la plantilla.
    pause
    exit /b 1
)

REM --- Generar la presentacion usando la plantilla limpia temporal ---
echo  [2/2] Generando presentacion...
"!PYTHON_EXE!" "%SCRIPT_DIR%Md2PPT.py" "%DOCS_DIR%\!MD_FILE!" "!PLANTILLA_TEMP!" "%DOCS_DIR%\!SALIDA!.pptx" --company "!COMPANY_NAME!"

REM --- Eliminar la plantilla temporal ---
if exist "!PLANTILLA_TEMP!" del "!PLANTILLA_TEMP!"

echo.
if %ERRORLEVEL%==0 (
    echo [OK] Proceso completado.
) else (
    echo [ERROR] El script finalizó con errores.
)
pause
