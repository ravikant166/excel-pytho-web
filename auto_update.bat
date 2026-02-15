@echo off
:: --- CONFIGURATION ---
:: Put the full path to your Excel file here
set INPUT_FILE="input file path"

:: Put the full path to where you want the HTML file saved
set OUTPUT_FILE="output file path"
:: ---------------------

echo Processing Excel to HTML...
echo Input: %INPUT_FILE%
echo Output: %OUTPUT_FILE%

:: This calls the python script and passes the paths as arguments
python excel_engine.py %INPUT_FILE% %OUTPUT_FILE%

echo.
echo Update Complete!
pause