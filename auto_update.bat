@echo off
:: --- CONFIGURATION ---
:: Put the full path to your Excel file here
set INPUT_FILE="C:\Users\ravik\Downloads\output\Project-Management-Sample-Data.xlsx"

:: Put the full path to where you want the HTML file saved
set OUTPUT_FILE="C:\Users\ravik\Downloads\output\Project-Management-Sample-Data.html"
:: ---------------------

echo Processing Excel to HTML...
echo Input: %INPUT_FILE%
echo Output: %OUTPUT_FILE%

:: This calls the python script and passes the paths as arguments
python excel_engine.py %INPUT_FILE% %OUTPUT_FILE%

echo.
echo Update Complete!
pause