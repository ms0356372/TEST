@echo off
cd /d "%~dp0"

echo Checking Python...
py --version
if errorlevel 1 goto NO_PYTHON

echo Installing requirements...
if exist requirements.txt py -m pip install -r requirements.txt

echo Installing PyInstaller...
py -m pip install pyinstaller

echo Cleaning old files...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo Building EXE...
py -m PyInstaller --onefile --noconsole --clean --name ExcelHealthTool main.py

if errorlevel 1 goto BUILD_ERROR

echo.
echo BUILD SUCCESS
echo EXE: dist\ExcelHealthTool.exe
echo.
pause
exit /b 0

:NO_PYTHON
echo.
echo Python launcher "py" was not found.
pause
exit /b 1

:BUILD_ERROR
echo.
echo BUILD FAILED
pause
exit /b 1
