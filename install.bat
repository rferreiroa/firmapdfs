@echo off
REM =============================================================================
REM FirmaPDFs - Installation Script for Windows
REM =============================================================================
REM This script creates a virtual environment and installs all dependencies.
REM Run as Administrator if you encounter permission issues.
REM =============================================================================

echo.
echo ============================================================
echo   FirmaPDFs - Installation Script
echo ============================================================
echo.

REM Check Python installation
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from https://python.org
    pause
    exit /b 1
)

echo [1/5] Checking Python version...
python --version

REM Check if venv exists
if exist "venv" (
    echo.
    echo Virtual environment already exists.
    set /p RECREATE="Do you want to recreate it? (y/N): "
    if /i "%RECREATE%"=="y" (
        echo Removing existing virtual environment...
        rmdir /s /q venv
    ) else (
        echo Using existing virtual environment.
        goto :install_deps
    )
)

echo.
echo [2/5] Creating virtual environment...
python -m venv venv
if errorlevel 1 (
    echo ERROR: Failed to create virtual environment
    pause
    exit /b 1
)

:install_deps
echo.
echo [3/5] Activating virtual environment...
call venv\Scripts\activate.bat

echo.
echo [4/5] Upgrading pip...
python -m pip install --upgrade pip

echo.
echo [5/5] Installing dependencies...
pip install -e ".[windows,dev]"

if errorlevel 1 (
    echo.
    echo WARNING: Some packages may have failed to install.
    echo Trying to install core dependencies only...
    pip install -e .
)

echo.
echo ============================================================
echo   Installation Complete!
echo ============================================================
echo.
echo To activate the virtual environment, run:
echo   venv\Scripts\activate.bat
echo.
echo To verify installation, run:
echo   firmapdfs --version
echo   firmapdfs status
echo.
echo To initialize project structure, run:
echo   firmapdfs init
echo.
pause
