@echo off
REM ─────────────────────────────────────────────────────────────────────────────
REM build_exe.bat  –  Build EO_Contract_Generator.exe (onedir, no console)
REM
REM Usage:  build_exe.bat
REM
REM Output: dist\EO_Contract_Generator\
REM   ├── EO_Contract_Generator.exe
REM   ├── _internal\          (runtime libs — keep together with exe)
REM   ├── data\               (copy your Excel here)
REM   ├── template\           (copy your .docx templates here)
REM   └── output\             (auto-created on first run)
REM ─────────────────────────────────────────────────────────────────────────────

setlocal

set DIST_DIR=dist\EO_Contract_Generator

echo [1/3] Building with PyInstaller ...
poetry run pyinstaller --noconfirm EO_Contract_Generator.spec
if errorlevel 1 (
    echo ERROR: PyInstaller build failed.
    exit /b 1
)

echo [2/3] Creating data\ and template\ placeholder folders in dist ...
if not exist "%DIST_DIR%\data"     mkdir "%DIST_DIR%\data"
if not exist "%DIST_DIR%\template" mkdir "%DIST_DIR%\template"
if not exist "%DIST_DIR%\output"   mkdir "%DIST_DIR%\output"

echo [3/3] Done.
echo.
echo Deployment folder: %DIST_DIR%\
echo.
echo Before running the exe, copy your files into the dist folder:
echo   %DIST_DIR%\data\E^&O summary table.xlsx
echo   %DIST_DIR%\template\Settlement Form_Draft.docx
echo   %DIST_DIR%\template\Settlement Form_Draft_Elan ^& SYNA.docx
echo.
endlocal
