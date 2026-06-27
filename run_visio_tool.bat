@echo off
REM ============================================================
REM  Double-click launcher for the Visio Text Replacer tool.
REM  Requires Python (https://www.python.org/downloads/).
REM ============================================================
setlocal
cd /d "%~dp0"

REM Prefer the Windows "py" launcher, then fall back to "python".
where py >nul 2>nul
if %errorlevel%==0 (
    py "%~dp0visio_replace_tool.py"
    goto :end
)

where python >nul 2>nul
if %errorlevel%==0 (
    python "%~dp0visio_replace_tool.py"
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
