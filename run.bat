@echo off
REM Interactive launcher. Requires Python 3.9+ on this PC.
REM On a PC without Python, use writeExcel.exe instead.
setlocal EnableExtensions
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 goto nopython

python write_excel.py
set "CODE=%ERRORLEVEL%"
if not "%CODE%"=="0" pause
exit /b %CODE%

:nopython
echo [ERROR] Python not found on this PC.
echo         Use writeExcel.exe (built by build_windows.bat) instead.
pause
exit /b 9009
