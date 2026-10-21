@echo off
setlocal EnableExtensions
cd /d "%~dp0"

echo ============================================================
echo  Building writeExcel.exe  (onefile, console)
echo ============================================================

where python >nul 2>nul
if errorlevel 1 goto nopython

echo [1/3] Installing dependencies ...
python -m pip install --upgrade pip >nul
python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 goto fail

echo [2/3] Running tests ...
python tests\selftest.py
if errorlevel 1 echo [WARN] self-test failed, building anyway ...

echo [3/3] Building exe ...
python -m PyInstaller --noconfirm --clean --onefile --name writeExcel --console --collect-submodules openpyxl write_excel.py
if errorlevel 1 goto fail

echo.
echo ============================================================
echo  Done.  Output:  dist\writeExcel.exe
echo  Copy that single file to any Windows PC - no Python needed.
echo  Keep fill.bat next to writeExcel.exe to call it from scripts.
echo ============================================================
pause
exit /b 0

:nopython
echo [ERROR] Python was not found.
echo         Install Python 3.9+ from https://www.python.org/downloads/
echo         and tick "Add python.exe to PATH" during setup.
pause
exit /b 1

:fail
echo.
echo [ERROR] Build failed. Scroll up to see the reason.
pause
exit /b 1
