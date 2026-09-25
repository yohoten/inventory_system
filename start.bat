@echo off
chcp 65001 >nul
title Inventory Management System v1.0
color 0A
cls

set PYTHON_SCRIPT=inventory_system.py
set LOG_FILE=error_log.txt
set MIN_PYTHON_VERSION=3.7

echo.             
echo  =============================
echo  Inventory Management System
echo  =============================
echo.

echo [1/4] Initializing startup environment...
timeout /t 1 /nobreak >nul

echo [2/4] Checking Python environment...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [Error] Python not found! Please install Python %MIN_PYTHON_VERSION%+
    echo Trying py command...
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [Error] Python still not found!
        echo Download: https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during installation
        echo.
        pause
        exit /b 1
    )
    set PYTHON_CMD=py
    echo [Info] Using py command
) else (
    set PYTHON_CMD=python
    echo [Info] Using python command
)

for /f "tokens=2" %%v in ('%PYTHON_CMD% --version 2^>^&1') do set PY_VER=%%v
echo [Info] Python version: %PY_VER%

echo [3/4] Verifying system files...
if not exist "%PYTHON_SCRIPT%" (
    echo.
    echo [Error] Core file missing: %PYTHON_SCRIPT%
    echo Current directory: %CD%
    echo.
    dir *.py
    echo.
    pause
    exit /b 2
)

echo [3.5/4] Checking dependencies...
%PYTHON_CMD% -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [Warning] Missing dependency: openpyxl
    echo Attempting auto-installation...
    %PYTHON_CMD% -m pip install openpyxl>=3.0.10
    if errorlevel 1 (
        echo [Error] Failed to install dependencies
        echo Please run manually: pip install openpyxl>=3.0.10
        echo.
        pause
        exit /b 3
    )
)

echo [4/4] Starting system...
echo ==================================================
echo.

echo [Debug] Current directory: %CD%
echo [Debug] Python command: %PYTHON_CMD%
echo [Debug] Script file: %PYTHON_SCRIPT%
echo.

%PYTHON_CMD% "%PYTHON_SCRIPT%" 2>>"%LOG_FILE%"
set EXIT_CODE=%errorlevel%

echo.
echo ==================================================
echo.

if "%EXIT_CODE%"=="0" (
    echo [Success] System exited normally
) else (
    echo [Warning] System exited with error code: %EXIT_CODE%
    echo [Info] Check log file: %LOG_FILE%
    if exist "%LOG_FILE%" (
        echo.
        echo ===== Error Log =====
        type "%LOG_FILE%"
        echo ====================
    )
)

echo.
echo Thank you for using! Press any key to close...
pause >nul
exit /b %EXIT_CODE%