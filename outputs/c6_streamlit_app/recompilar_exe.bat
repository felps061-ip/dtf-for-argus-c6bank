@echo off
setlocal

cd /d "%~dp0"

set "APP_NAME=Auto C6BANK - Datafast for Argus"
set "OUT_DIR=%~dp0..\c6_executavel"
set "OUT_EXE=%OUT_DIR%\%APP_NAME%.exe"
set "STAMP=%DATE:/=-%_%TIME::=-%"
set "STAMP=%STAMP: =0%"
set "BUILD_ROOT=%LOCALAPPDATA%\Temp\AutoC6Build\%STAMP%"
set "BUILD_DIR=%BUILD_ROOT%\build"
set "DIST_DIR=%BUILD_ROOT%\dist"
set "SPEC_DIR=%BUILD_ROOT%\spec"
set "PYINSTALLER_CONFIG_DIR=%BUILD_ROOT%\cache"
set "TEMP=%BUILD_ROOT%\tmp"
set "TMP=%BUILD_ROOT%\tmp"

echo.
echo Recompilando %APP_NAME%
echo Pasta: %CD%
echo Build temporario: %BUILD_ROOT%
echo.

where py >nul 2>nul
if %errorlevel%==0 (
    py -3.13 -c "import sys" >nul 2>nul && set "PY_CMD=py -3.13"
    if not defined PY_CMD py -3.12 -c "import sys" >nul 2>nul && set "PY_CMD=py -3.12"
)

if not defined PY_CMD (
    python -c "import sys" >nul 2>nul && set "PY_CMD=python"
)

if not defined PY_CMD (
    echo Nao encontrei Python instalado.
    echo Instale Python 3.12 ou 3.13 e rode este arquivo de novo.
    pause
    exit /b 1
)

%PY_CMD% -c "import sys; raise SystemExit(0 if sys.version_info < (3, 14) else 1)"
if errorlevel 1 (
    echo.
    echo Seu Python atual parece ser 3.14.
    echo Se o PyInstaller falhar, instale Python 3.12 ou 3.13 pelo site python.org
    echo e rode este arquivo novamente.
    echo.
)

if not exist ".venv_build\Scripts\python.exe" (
    %PY_CMD% -m venv .venv_build
    if errorlevel 1 goto erro
)

call ".venv_build\Scripts\activate.bat"
if errorlevel 1 goto erro

python -m pip install --upgrade pip
if errorlevel 1 goto erro

python -m pip install pyinstaller pandas XlsxWriter openpyxl xlrd
if errorlevel 1 goto erro

mkdir "%BUILD_DIR%" "%DIST_DIR%" "%SPEC_DIR%" "%PYINSTALLER_CONFIG_DIR%" "%TEMP%" >nul 2>nul

python -m PyInstaller ^
    --noconfirm ^
    --onefile ^
    --windowed ^
    --name "%APP_NAME%" ^
    --workpath "%BUILD_DIR%" ^
    --distpath "%DIST_DIR%" ^
    --specpath "%SPEC_DIR%" ^
    desktop_app.py
if errorlevel 1 goto erro

if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"
taskkill /IM "%APP_NAME%.exe" /F >nul 2>nul
copy /Y "%DIST_DIR%\%APP_NAME%.exe" "%OUT_EXE%"
if errorlevel 1 goto erro

echo.
echo Pronto. Executavel atualizado em:
echo %OUT_EXE%
echo.
pause
exit /b 0

:erro
echo.
echo Falha ao recompilar. Veja a mensagem acima.
echo.
pause
exit /b 1
