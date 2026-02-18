@echo off
chcp 65001 >nul
cd /d "%~dp0"

set TASK_NAME=FIN_Проверка_разметки
set BAT_PATH=%~dp0ПроверитьРазметку_авто.bat

echo Установка ежедневной задачи "%TASK_NAME%"
echo Запуск: каждый день в 13:45 (Москва)
echo Скрипт: %BAT_PATH%
echo.

schtasks /create /tn "%TASK_NAME%" /tr "\"%BAT_PATH%\"" /sc daily /st 13:45 /f

if %errorlevel% equ 0 (
    echo.
    echo Задача создана успешно.
    echo Проверить: Планировщик заданий Windows - "%TASK_NAME%"
) else (
    echo.
    echo Ошибка создания задачи. Запустите от имени администратора.
)

pause
