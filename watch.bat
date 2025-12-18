@echo off
REM =============================================================================
REM FirmaPDFs - Watch Mode
REM =============================================================================
REM Starts the file watcher to automatically process new documents
REM Press Ctrl+C to stop
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

echo Starting FirmaPDFs in watch mode...
echo Press Ctrl+C to stop.
echo.

firmapdfs watch %*
