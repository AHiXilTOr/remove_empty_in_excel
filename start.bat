@echo off
chcp 65001 > nul
set PYTHON_SCRIPT_PATH=main.py
set VIRTUAL_ENV_PATH=venv

if not exist %VIRTUAL_ENV_PATH% (
    echo Виртуальное окружение не найдено. Активируйте его перед выполнением скрипта.
    pause
    exit /b
)

call %~dp0\%VIRTUAL_ENV_PATH%\Scripts\activate

mode con cols=120 lines=40

start /MAX python %~dp0\%PYTHON_SCRIPT_PATH%

pause