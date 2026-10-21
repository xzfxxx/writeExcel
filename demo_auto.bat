@echo off
REM Demo: run the whole pipeline automatically, no manual steps.
setlocal EnableExtensions
cd /d "%~dp0"

if exist "example\auto-output.xlsx" del /q "example\auto-output.xlsx"

call "%~dp0fill.bat" "%~dp0example\template-sample.xlsx" "%~dp0example\data-sample-multi.txt" "%~dp0example\auto-output.xlsx"
set "CODE=%ERRORLEVEL%"

echo.
if "%CODE%"=="0" goto ok
echo FAILED - exit code %CODE%
goto end

:ok
echo OK - generated: example\auto-output.xlsx

:end
pause
exit /b %CODE%
