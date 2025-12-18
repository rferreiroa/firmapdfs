@echo off
REM =============================================================================
REM FirmaPDFs - Run Script
REM =============================================================================
REM Quick script to run FirmaPDFs with virtual environment activated
REM =============================================================================

REM Activate virtual environment
if exist "venv\Scripts\activate.bat" (
    call venv\Scripts\activate.bat
) else (
    echo ERROR: Virtual environment not found.
    echo Please run install.bat first.
    pause
    exit /b 1
)

REM Run FirmaPDFs with all arguments passed to this script
firmapdfs %*
