@echo off
REM 自动调用示例：等价于 fill.bat 模板 数据 输出
REM 用于在 Windows 上点两下验证整条链路（无任何手动操作）
setlocal
cd /d "%~dp0"

if exist "example\自动输出.xlsx" del /q "example\自动输出.xlsx"

call "%~dp0fill.bat" "%~dp0example\模板示例.xlsx" "%~dp0example\数据示例_多表.txt" "%~dp0example\自动输出.xlsx"
set "CODE=%ERRORLEVEL%"

echo.
echo ============================================================
echo  Exit code: %CODE%
if "%CODE%"=="0" (
    echo  OK - generated: example\自动输出.xlsx
) else (
    echo  FAILED - see message above
)
echo ============================================================
pause
exit /b %CODE%
