@echo off
setlocal
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto usepython
py -3 -m venv .venv
if errorlevel 1 goto fail
goto deps
:usepython
python -m venv .venv
if errorlevel 1 goto fail
:deps
.venv\Scripts\python.exe -m ensurepip --upgrade
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
echo Installation complete. Run start.bat.
pause
exit /b 0
:fail
echo Installation failed. Install Python 3.11 or later with Tcl/Tk, then retry.
pause
exit /b 1
