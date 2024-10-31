@echo off
setlocal enabledelayedexpansion

cd /d %~dp0

echo Checking Python installation...
where python >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Python is not installed or not in PATH
    echo Please install Python and try again
    pause
    exit /b 1
)

echo Checking for input_data.xlsx...
if not exist "input_data.xlsx" (
    echo Error: input_data.xlsx not found in the current directory
    echo Please ensure your Excel file is named 'input_data.xlsx' and is in the same folder as this script
    pause
    exit /b 1
)

echo Checking required Python packages...
python -c "import pandas" 2>nul
if %errorlevel% neq 0 (
    echo Required packages are not installed
    echo Please run setup.bat first to install dependencies
    pause
    exit /b 1
)

echo.
echo Starting Product Validator...
echo.
echo Note: This tool will validate:
echo - Product variant ordering and titles
echo - Image dimensions ^(825x825 or 1:1 ratio^)
echo - Price hierarchy and relationships
echo - Inventory quantities
echo - HTML content
echo.

python product_validator.py "input_data.xlsx"
if %errorlevel% neq 0 (
    echo.
    echo Error: The validator encountered an error
    echo Please check the error message above
    pause
    exit /b 1
)

echo.
echo Script execution completed. Press any key to exit...
pause > nul