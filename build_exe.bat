@echo off
REM ============================================================
REM  Builds a standalone Windows EXE of Drawing & BOM Studio.
REM  Run this ONCE; afterwards you can share/run the EXE without
REM  needing Python installed. (Building still needs Python.)
REM ============================================================
setlocal
cd /d "%~dp0"

REM Pick the Python command that exists (py launcher or python).
set "PY=py"
where py >nul 2>nul || set "PY=python"
where %PY% >nul 2>nul || (
    echo.
    echo  Python was not found. Install it from
    echo  https://www.python.org/downloads/  (tick "Add to PATH"^),
    echo  then run this file again.
    echo.
    pause
    exit /b 1
)

echo Updating pip...
%PY% -m pip install --upgrade pip setuptools wheel >nul 2>nul

REM --- PyInstaller is REQUIRED to build the EXE. -------------------------
echo Installing PyInstaller...
%PY% -m pip install --upgrade pyinstaller || (
    echo.
    echo  Could not install PyInstaller. Check your internet connection and
    echo  that pip is up to date, then run this file again.
    pause
    exit /b 1
)

REM --- Optional dependencies: install each on its own so one that has no
REM     build for your Python version can't block the others. ------------
echo Installing pypdf (for "Split sheets into PDFs")...
%PY% -m pip install --upgrade pypdf || echo   NOTE: pypdf did not install -- the "Split sheets to PDFs" feature will be unavailable.

echo Installing Pillow (for the robot mascot art)...
%PY% -m pip install --upgrade pillow || echo   NOTE: Pillow did not install -- the mascot art will be skipped.

echo Installing tkinterdnd2 (for drag-and-drop)...
set "DND="
%PY% -m pip install --upgrade tkinterdnd2 >nul 2>nul && %PY% -c "import tkinterdnd2" >nul 2>nul && set "DND=1"
if defined DND (
    echo   Drag-and-drop: enabled.
) else (
    echo   NOTE: tkinterdnd2 has no build for this Python version -- building
    echo         without drag-and-drop. The "Add file..." buttons still work.
)

REM Only ask PyInstaller to bundle tkinterdnd2 when it is actually installed;
REM otherwise "--collect-all tkinterdnd2" would fail the build.
set "COLLECT="
if defined DND set "COLLECT=--collect-all tkinterdnd2"

echo.
echo Building the EXE...
%PY% -m PyInstaller --onefile --windowed --name DrawingBOMStudio --icon robot_icon.ico %COLLECT% drawing_bom_studio.py || (
    echo Build failed.
    pause
    exit /b 1
)

echo.
echo  ------------------------------------------------------------
echo   Done!  Your program is here:
echo       dist\DrawingBOMStudio.exe
echo.
echo   You can move that EXE anywhere and double-click to run it.
echo   (PDF export still needs LibreOffice installed.)
echo  ------------------------------------------------------------
echo.
pause
endlocal
