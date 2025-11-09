@echo off
REM Bank Statement Categorization Tool
REM Double-click to categorize your bank statements

title Bank Statement Categorization

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

REM Run the categorization tool
python categorize_statement.py

