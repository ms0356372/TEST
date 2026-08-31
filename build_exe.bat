@echo off
setlocal EnableExtensions EnableDelayedExpansion
chcp 65001 >nul
cd /d "%~dp0"
set "LOG=build_log.txt"
set "RESULT=%TEMP%\excel_tool_self_test.txt"
set "EXE=dist\Excel整理分析工具.exe"
set "RELEASE=release\Excel整理分析工具_v1.0.0.exe"
>"%LOG%" echo Build started: %date% %time%
echo ========================================
echo Excel整理分析工具 - EXE 自動建置程式
echo ========================================
echo [1/9] 檢查 Python...
python --version >>"%LOG%" 2>&1 || (echo [失敗] 找不到 Python 3.13 或相容版本。& goto :fail)
python --version
echo [2/9] 檢查 pip...
python -m pip --version >>"%LOG%" 2>&1 || (echo [失敗] 找不到 pip。& goto :fail)
echo [成功] pip 正常
echo [3/9] 安裝/確認必要套件...
python -m pip install -r requirements.txt >>"%LOG%" 2>&1 || (echo [失敗] 套件安裝失敗。& goto :fail)
python -m PyInstaller --version >>"%LOG%" 2>&1
echo [成功] openpyxl / PyInstaller
echo [4/9] 執行核心單元測試...
python test_transform.py >>"%LOG%" 2>&1 || (echo [失敗] 核心轉換測試未通過。& goto :fail)
echo [成功] 核心轉換測試通過
echo [5/9] 執行 Python Self Test...
python main.py --self-test >>"%LOG%" 2>&1 || (echo [失敗] Python Self Test 未通過。& goto :fail)
echo [成功] Python Self Test
echo [6/9] 清除舊版 Build...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "Excel整理分析工具.spec" del /q "Excel整理分析工具.spec"
echo [成功] 清除完成
echo [7/9] 建立 EXE...
python -m PyInstaller --onefile --windowed --clean --name "Excel整理分析工具" main.py >>"%LOG%" 2>&1 || (echo [失敗] PyInstaller 建置失敗。& goto :fail)
echo [成功] PyInstaller
echo [8/9] 檢查 EXE...
if not exist "%EXE%" (echo [失敗] EXE 建立失敗。& goto :fail)
echo [成功] %EXE%
echo [9/9] 執行 EXE Self Test...
if exist "%RESULT%" del /q "%RESULT%"
"%EXE%" --self-test --result-file "%RESULT%"
set /a WAIT=0
:wait_result
if exist "%RESULT%" goto result_ready
ping 127.0.0.1 -n 2 >nul
set /a WAIT+=1
if !WAIT! GEQ 60 (echo [失敗] EXE Self Test 逾時。& goto :fail)
goto wait_result
:result_ready
type "%RESULT%" >>"%LOG%"
findstr /b /c:"PASS" "%RESULT%" >nul || (echo [失敗] EXE Self Test 未通過。& goto :fail)
if not exist release mkdir release
copy /y "%EXE%" "%RELEASE%" >>"%LOG%"
for %%A in ("%RELEASE%") do set SIZE=%%~zA
>>"%LOG%" echo EXE: %CD%\%RELEASE%
>>"%LOG%" echo EXE bytes: !SIZE!
echo ========================================
echo BUILD SUCCESS
echo Python Test : PASS
echo Core Test   : PASS
echo Self Test   : PASS
echo PyInstaller : PASS
echo EXE Test    : PASS
echo 正式 EXE：%CD%\%RELEASE%
echo EXE 大小：!SIZE! bytes
pause
exit /b 0
:fail
>>"%LOG%" echo BUILD FAILED: %date% %time%
echo ========================================
echo BUILD FAILED
echo 詳細資訊請查看 %LOG%
pause
exit /b 1
