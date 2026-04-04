@echo off
setlocal
call "%~dp0setup_windows.bat" %*
exit /b %errorlevel%
