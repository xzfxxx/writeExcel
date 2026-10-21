@echo off
REM 双击运行：进入交互模式（可以把 Excel 模板和 txt 直接拖进窗口）
REM 需要本机已安装 Python 3.9+；没有 Python 的电脑请用 writeExcel.exe
setlocal
cd /d "%~dp0"
chcp 65001 >nul

where python >nul 2>nul
if errorlevel 1 (
    echo [错误] 本机没有找到 Python。
    echo        如果你在"白板"电脑上使用，请改用 build_windows.bat 打包出来的 writeExcel.exe。
    echo.
    pause
    exit /b 1
)

python write_excel.py
set EXITCODE=%ERRORLEVEL%
if not "%EXITCODE%"=="0" pause
exit /b %EXITCODE%
