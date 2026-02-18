@echo off
chcp 65001 >nul
cd /d "%~dp0"
python update_reestr_from_51.py
pause