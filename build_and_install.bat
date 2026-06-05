@echo off
chcp 65001 >nul
title Time Tracker - Compilation

cd /d "%~dp0"

echo.
echo  ╔══════════════════════════════════════╗
echo  ║     Time Tracker - Compilation       ║
echo  ╚══════════════════════════════════════╝
echo.

python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] Python n'est pas installe ou pas dans le PATH.
    pause
    exit /b 1
)
echo  [OK] Python detecte

if not exist "main.py" (
    echo  [ERREUR] Fichier source introuvable : main.py
    pause
    exit /b 1
)

echo  [..] Installation des dependances...
pip install -r requirements.txt pyinstaller --quiet --disable-pip-version-check
if errorlevel 1 (
    echo  [ERREUR] Impossible d'installer les dependances.
    pause
    exit /b 1
)
echo  [OK] Dependances installees

echo  [..] Compilation en .exe (peut prendre 1-2 minutes)...
set "BUILD_LOG=%TEMP%\timetracker_build.log"
pyinstaller --onefile --windowed --name "TimeTracker" --clean main.py >"%BUILD_LOG%" 2>&1
if errorlevel 1 (
    echo  [ERREUR] La compilation a echoue.
    type "%BUILD_LOG%"
    pause
    exit /b 1
)

if not exist "dist\TimeTracker.exe" (
    echo  [ERREUR] dist\TimeTracker.exe introuvable apres compilation.
    type "%BUILD_LOG%"
    pause
    exit /b 1
)
echo  [OK] .exe compile : dist\TimeTracker.exe

echo.
echo  ════════════════════════════════════════
echo  Compilation terminee !
echo.
echo  Double-clique sur dist\TimeTracker.exe
echo  pour installer l'app sur ton PC.
echo  ════════════════════════════════════════
echo.

set /p LAUNCH=" Lancer dist\TimeTracker.exe maintenant ? (O/N) : "
if /i "%LAUNCH%"=="O" start "" "%~dp0dist\TimeTracker.exe"

pause
