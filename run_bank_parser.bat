@echo off
setlocal
set "ROOT=%~dp0"
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"

if not exist "%VENV_PY%" goto system_python
"%VENV_PY%" --version >nul 2>nul
if errorlevel 1 goto system_python
"%VENV_PY%" "%ROOT%run.py" %*
exit /b %errorlevel%

:system_python
py -3 --version >nul 2>nul
if errorlevel 1 goto path_python
py -3 "%ROOT%run.py" %*
exit /b %errorlevel%

:path_python
python "%ROOT%run.py" %*
exit /b %errorlevel%
