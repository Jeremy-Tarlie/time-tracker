@echo off
chcp 65001 >nul
title Time Tracker - Installation

echo.
echo  ╔══════════════════════════════════════╗
echo  ║     Time Tracker - Installation      ║
echo  ╚══════════════════════════════════════╝
echo.

:: ── 1. Vérifier Python ──────────────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] Python n'est pas installe ou pas dans le PATH.
    echo  Telecharge-le sur https://www.python.org/downloads/
    echo  /!\ Coche bien "Add Python to PATH" lors de l'installation.
    pause
    exit /b 1
)
echo  [OK] Python detecte

:: ── 2. Installer pyinstaller + openpyxl ─────────────────────────────────────
echo  [..] Installation des dependances...
pip install pyinstaller openpyxl --quiet --disable-pip-version-check
if errorlevel 1 (
    echo  [ERREUR] Impossible d'installer les dependances.
    pause
    exit /b 1
)
echo  [OK] Dependances installees

:: ── 3. Compiler le .exe ──────────────────────────────────────────────────────
echo  [..] Compilation en .exe (peut prendre 1-2 minutes)...
pyinstaller --onefile --windowed --name "TimeTracker" --clean time_tracker.py >nul 2>&1
if errorlevel 1 (
    echo  [ERREUR] La compilation a echoue.
    pause
    exit /b 1
)
echo  [OK] .exe compile avec succes

:: ── 4. Copier l'exe dans un dossier permanent ────────────────────────────────
set "INSTALL_DIR=%USERPROFILE%\AppData\Local\TimeTracker"
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

copy /Y "dist\TimeTracker.exe" "%INSTALL_DIR%\TimeTracker.exe" >nul
echo  [OK] TimeTracker.exe copie dans : %INSTALL_DIR%

:: ── 5. Ajouter au démarrage Windows (registre) ──────────────────────────────
reg add "HKCU\Software\Microsoft\Windows\CurrentVersion\Run" ^
    /v "TimeTracker" ^
    /t REG_SZ ^
    /d "\"%INSTALL_DIR%\TimeTracker.exe\"" ^
    /f >nul

if errorlevel 1 (
    echo  [ERREUR] Impossible d'ajouter au demarrage.
    pause
    exit /b 1
)
echo  [OK] Ajout au demarrage Windows

:: ── 6. Lancer l'app maintenant ───────────────────────────────────────────────
echo.
echo  ════════════════════════════════════════
echo  Installation terminee avec succes !
echo.
echo  L'app se lancera automatiquement
echo  a chaque demarrage de Windows.
echo  ════════════════════════════════════════
echo.

set /p LAUNCH=" Lancer Time Tracker maintenant ? (O/N) : "
if /i "%LAUNCH%"=="O" start "" "%INSTALL_DIR%\TimeTracker.exe"

echo.
echo  Pour desinstaller : supprimer %INSTALL_DIR%
echo  et effacer la cle "TimeTracker" dans :
echo  HKCU\Software\Microsoft\Windows\CurrentVersion\Run
echo.
pause