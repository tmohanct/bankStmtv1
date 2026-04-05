@echo off
setlocal
set "ROOT=%~dp0"

py -3 "%ROOT%build_fresh_machine_package.py" %* 2>nul
set "EXITCODE=%errorlevel%"
if not "%EXITCODE%"=="9009" exit /b %EXITCODE%

python "%ROOT%build_fresh_machine_package.py" %*
exit /b %errorlevel%
