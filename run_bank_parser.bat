@echo off
setlocal
set "ROOT=%~dp0"
set "VENV_PY=%ROOT%.venv\Scripts\python.exe"

if exist "%VENV_PY%" (
    "%VENV_PY%" "%ROOT%src\code\run.py" %*
    exit /b %errorlevel%
)

py -3 "%ROOT%src\code\run.py" %* 2>nul
set "EXITCODE=%errorlevel%"
if not "%EXITCODE%"=="9009" exit /b %EXITCODE%

python "%ROOT%src\code\run.py" %*
exit /b %errorlevel%
