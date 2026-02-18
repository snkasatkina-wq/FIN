@echo off
chcp 65001 >nul
cd /d "%~dp0"

if not exist "%~dp0logs" mkdir "%~dp0logs"
set LOGFILE=%~dp0logs\unmarked.log

echo. >> "%LOGFILE%"
echo ========== %date% %time% ========== >> "%LOGFILE%"
python unmarked_alert.py >> "%LOGFILE%" 2>&1
echo Завершено. >> "%LOGFILE%"
