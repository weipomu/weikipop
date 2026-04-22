@echo off
cd /d "%~dp0"
title Weikipop Builder

echo ============================================================
echo  Weikipop Build Script
echo ============================================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH.
    echo Install Python 3.11-3.13 from https://python.org
    echo Make sure to tick "Add Python to PATH" during install.
    echo.
    pause & exit /b 1
)
python --version

:: Install dependencies
echo.
echo Installing dependencies...
pip install pyinstaller
pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: Failed to install one or more dependencies.
    echo See the output above for details.
    echo.
    pause & exit /b 1
)

:: Copy JPDB frequency zip if present
if exist "_Freq__JPDB.zip" (
    echo.
    echo Found _Freq__JPDB.zip -- copying to data\ for better frequency sorting...
    if not exist "data" mkdir data
    copy /Y "_Freq__JPDB.zip" "data\_Freq__JPDB.zip" >nul
) else if exist "data\_Freq__JPDB.zip" (
    echo.
    echo Found data\_Freq__JPDB.zip -- JPDB frequencies will be baked in.
)

:: Build dictionary
echo.
echo ============================================================
echo  Building dictionary.pkl -- downloads ~30 MB, takes ~1-2 min
echo  Do not close this window.
echo ============================================================
echo.
python -m scripts.build_dictionary
if errorlevel 1 (
    echo.
    echo ERROR: Dictionary build failed. See the output above.
    echo.
    pause & exit /b 1
)

:: Run PyInstaller
echo.
echo ============================================================
echo  Building weikipop.exe...
echo ============================================================
echo.
pyinstaller weikipop_win_x64.spec --noconfirm
if errorlevel 1 (
    echo.
    echo ERROR: PyInstaller build failed. See the output above.
    echo.
    pause & exit /b 1
)

:: Copy required files
echo.
echo Copying files to dist\...
if not exist "dist" mkdir dist
copy /Y "dictionary.pkl" "dist\dictionary.pkl" >nul
if exist "config.ini" copy /Y "config.ini" "dist\config.ini" >nul

:: Refresh icon cache
taskkill /f /im explorer.exe >nul 2>&1
del /f /q "%localappdata%\IconCache.db" >nul 2>&1
del /f /q "%localappdata%\Microsoft\Windows\Explorer\iconcache*" >nul 2>&1
start explorer.exe

echo.
echo ============================================================
echo  Done! dist\ contains weikipop.exe and dictionary.pkl
echo ============================================================
echo.
pause
