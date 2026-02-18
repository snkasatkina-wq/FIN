@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo Проверка нерозмеченных операций и формирование сообщения...
python unmarked_alert.py
pause
