@echo off
setlocal
py -3 -V >nul 2>&1
if not errorlevel 1 (
    py -3 "%~dp0stability96_web.py"
) else (
    python "%~dp0stability96_web.py"
)
pause
