@echo off
title ICAI NCE Financial Statement Converter - ASAR ^& CO
echo.
echo  ====================================================
echo   ICAI NCE Financial Statement Converter
echo   ASAR ^& CO Chartered Accountants
echo  ====================================================
echo.
cd /d "%~dp0"
pip install -r requirements.txt --quiet 2>nul
echo  Starting server...
echo.
python app.py
pause
