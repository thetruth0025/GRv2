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

echo Installing PyInstaller (if not already present)...
%PY% -m pip install --upgrade pyinstaller pillow tkinterdnd2 || (
    echo Failed to install PyInstaller. Check your internet connection.
    pause
    exit /b 1
)

echo.
echo Building the EXE...
%PY% -m PyInstaller --onefile --windowed --name DrawingBOMStudio --icon robot_icon.ico --collect-all tkinterdnd2 drawing_bom_studio.py || (
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
