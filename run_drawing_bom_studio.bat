@echo off
REM ============================================================
REM  Double-click launcher for Drawing & BOM Studio.
REM  Requires Python (https://www.python.org/downloads/).
REM ============================================================
setlocal
cd /d "%~dp0"

REM Optional extras, installed one at a time so a package with no build for
REM your Python version (often tkinterdnd2 on a brand-new Python) can't stop
REM the others from installing. pypdf powers "Split sheets into PDFs"; Pillow
REM draws the mascot; tkinterdnd2 adds drag-and-drop. All are optional.
REM Prefer the Windows "py" launcher, then fall back to "python".
where py >nul 2>nul
if %errorlevel%==0 (
    py -m pip install --quiet --disable-pip-version-check pypdf >nul 2>nul
    py -m pip install --quiet --disable-pip-version-check pillow >nul 2>nul
    py -m pip install --quiet --disable-pip-version-check tkinterdnd2 >nul 2>nul
    py "%~dp0drawing_bom_studio.py"
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python -m pip install --quiet --disable-pip-version-check pypdf >nul 2>nul
    python -m pip install --quiet --disable-pip-version-check pillow >nul 2>nul
    python -m pip install --quiet --disable-pip-version-check tkinterdnd2 >nul 2>nul
    python "%~dp0drawing_bom_studio.py"
    goto :end
)

echo.
echo  ------------------------------------------------------------
echo   Python was not found on this computer.
echo.
echo   1. Install it from:  https://www.python.org/downloads/
echo   2. During setup, TICK the box "Add python.exe to PATH".
echo   3. Then double-click this file again.
echo  ------------------------------------------------------------
echo.
pause

:end
endlocal
