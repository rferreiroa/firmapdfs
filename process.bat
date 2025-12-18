@echo off
REM =============================================================================
REM FirmaPDFs - Process Documents
REM =============================================================================
REM Quick script to process all documents in the inbox folder
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

echo Processing documents from inbox...
echo.

REM Process all documents
firmapdfs process %*

echo.
echo Done!
pause
