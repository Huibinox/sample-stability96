@echo off
cd /d "%~dp0"
set "PYTHON_EXE=%LOCALAPPDATA%\Python\pythoncore-3.14-64\pythonw.exe"
if exist "%PYTHON_EXE%" (
  start "" "%PYTHON_EXE%" "%~dp0stability96.py"
  exit /b
)
where pyw >nul 2>nul
if %errorlevel%==0 (
  start "" pyw -3 "%~dp0stability96.py"
  exit /b
)
echo Python 3.10+ with tkinter and openpyxl is required.
pause
