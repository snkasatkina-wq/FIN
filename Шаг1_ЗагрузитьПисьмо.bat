@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Загрузка письма от vtb-inform из mail.ru...
python step1_fetch_email.py
pause
