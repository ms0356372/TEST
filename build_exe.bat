@echo off
setlocal EnableExtensions
cd /d "%~dp0"
chcp 65001 >nul 2>nul
color 07
cls
echo ========================================
echo Excel Transform Tool - Windows EXE Build
echo ========================================
echo.

set "PYTHON_CMD="
where py >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3"
if not defined PYTHON_CMD (
    where python >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo [FAILED] Python was not found.
    echo Please install Python 3.13 or a compatible version.
    echo.
    pause
    exit /b 1
)

%PYTHON_CMD% build_windows.py
set "BUILD_EXIT=%ERRORLEVEL%"
echo.
if not "%BUILD_EXIT%"=="0" (
    echo ========================================
    echo BUILD FAILED - see build_log.txt
    echo ========================================
) else (
    echo ========================================
    echo BUILD SUCCESS
    echo ========================================
)
echo.
pause
exit /b %BUILD_EXIT%
