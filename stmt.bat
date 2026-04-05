@echo off
setlocal
set "ROOT=%~dp0"

py -3 "%ROOT%src\code\run.py" %* 2>nul
set "EXITCODE=%errorlevel%"
if not "%EXITCODE%"=="9009" exit /b %EXITCODE%

python "%ROOT%src\code\run.py" %*
exit /b %errorlevel%
