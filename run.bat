@echo off
cd /d %~dp0

echo Starting Product Validator...
echo.
echo Note: This tool will validate:
echo - Product variant ordering and titles
echo - Image dimensions (825x825 or 1:1 ratio)
echo - Price hierarchy and relationships
echo - Inventory quantities
echo - HTML content
echo.

python product_validator.py

echo.
echo Script execution completed. Press any key to exit...
pause > nul