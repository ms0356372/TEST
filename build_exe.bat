@echo off
setlocal
cd /d "%~dp0"
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
python -m PyInstaller --onefile --noconsole --clean --name "Excel健檢分析工具" main.py
if errorlevel 1 (echo 打包失敗 & pause & exit /b 1)
echo 完成：dist\Excel健檢分析工具.exe
pause
