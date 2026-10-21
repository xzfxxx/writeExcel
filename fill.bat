@echo off
REM ============================================================
REM  fill.bat - call writeExcel.exe automatically (no interaction)
REM
REM  Usage:
REM    fill.bat ^<template.xlsx^> ^<data.txt^> [output.xlsx] [options...]
REM
REM  arg1  : Excel template path
REM  arg2  : data txt path
REM  arg3  : output file name or full path
REM          (omitted = default "<template>_filled.xlsx" next to template)
REM  arg4+ : forwarded to writeExcel.exe, e.g. -y  --parse-numbers
REM
REM  Exit codes: 0=ok  2=bad usage  3=data error  4=write error
REM              9009=writeExcel.exe not found
REM
REM  Notes:
REM    * stdin is redirected from NUL, so the program can never
REM      wait for a keypress and hang.
REM    * keep this file using CRLF line endings (Windows batch).
REM ============================================================
setlocal EnableExtensions

set "EXE=%~dp0writeExcel.exe"

if not exist "%EXE%" goto noexe
if "%~1"=="" goto usage
if "%~2"=="" goto usage
if "%~3"=="" goto no_output

"%EXE%" "%~1" "%~2" "%~3" --no-pause %4 %5 %6 %7 %8 %9 <nul
exit /b %ERRORLEVEL%

:no_output
"%EXE%" "%~1" "%~2" --no-pause %4 %5 %6 %7 %8 %9 <nul
exit /b %ERRORLEVEL%

:noexe
echo [ERROR] writeExcel.exe not found in: %~dp0
echo         Put writeExcel.exe and fill.bat in the same folder.
exit /b 9009

:usage
echo Usage: fill.bat ^<template.xlsx^> ^<data.txt^> [output.xlsx] [options...]
echo Examples:
echo   fill.bat C:\tpl\template.xlsx C:\data\data.txt
echo   fill.bat C:\tpl\template.xlsx C:\data\data.txt C:\out\result.xlsx -y
exit /b 2
