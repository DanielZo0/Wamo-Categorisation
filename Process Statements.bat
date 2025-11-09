@echo off
REM Bank Statement Batch Processor
REM Double-click to process your bank statements

title Statement Processor

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.8 or higher from python.org
    echo.
    pause
    exit /b 1
)

REM Run the batch processor
python batch_statement_processor.py

