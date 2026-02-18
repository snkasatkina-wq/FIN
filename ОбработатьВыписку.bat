@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Запуск обработки выписки из банка...
python main.py
pause
