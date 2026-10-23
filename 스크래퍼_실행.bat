@echo off
set PYTHONIOENCODING=utf-8
chcp 65001 > nul
cd /d "%~dp0"
echo ===================================================
echo   rERP Purchasing Scraper Starting...
echo ===================================================
python local_tools\scraper.py
pause
