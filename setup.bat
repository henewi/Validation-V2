@echo off
cd /d %~dp0

echo Setting up environment for Product Validator...
echo.

REM Install wheel and setuptools first
python -m pip install --upgrade pip
python -m pip install --upgrade wheel
python -m pip install --upgrade setuptools

echo.
echo Installing required packages...
echo.

REM Install main dependencies
pip install pandas
pip install openpyxl
pip install requests
pip install urllib3
pip install Pillow
pip install beautifulsoup4
pip install html5lib

echo.
echo Setup complete! You can now use run.bat to execute the validator.
echo.
pause