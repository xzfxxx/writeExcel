@echo off
REM ============================================================
REM  fill.bat - 调用 writeExcel.exe 自动填充 Excel 模板
REM
REM  用法:
REM    fill.bat ^<template.xlsx^> ^<data.txt^> [output.xlsx] [附加参数...]
REM
REM  参数1: Excel 模板路径
REM  参数2: 数据 txt 路径
REM  参数3: 输出文件名/路径（不传 = 默认生成 "模板名_已填充.xlsx"）
REM  之后可跟任意 writeExcel 参数，例如 -y（允许覆盖）、--parse-numbers
REM
REM  特点: 全程无人交互，执行完立即返回；退出码 0 = 成功
REM ============================================================
setlocal EnableExtensions

set "EXE=%~dp0writeExcel.exe"
if not exist "%EXE%" (
    echo [ERROR] writeExcel.exe not found next to this script: "%~dp0"
    echo         Put fill.bat and writeExcel.exe in the same folder.
    exit /b 9009
)

if "%~1"=="" goto usage
if "%~2"=="" goto usage

if "%~3"=="" (
    "%EXE%" "%~1" "%~2" --no-pause
) else (
    "%EXE%" "%~1" "%~2" "%~3" --no-pause %4 %5 %6 %7 %8 %9
)

exit /b %ERRORLEVEL%

:usage
echo Usage:
echo   fill.bat ^<template.xlsx^> ^<data.txt^> [output.xlsx] [options...]
echo.
echo Examples:
echo   fill.bat C:\tpl\模板.xlsx C:\data\数据.txt
echo   fill.bat C:\tpl\模板.xlsx C:\data\数据.txt C:\out\结果.xlsx -y
exit /b 2
