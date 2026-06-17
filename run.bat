@echo off
chcp 65001 >nul
cd /d "%~dp0"

title RenSer Letter Bot

if not exist ".venv\Scripts\python.exe" (
    echo [1/2] Создаю виртуальное окружение...
    python -m venv .venv
    if errorlevel 1 (
        echo Ошибка: не найден Python. Установите Python 3.10+ с python.org
        pause
        exit /b 1
    )
    echo [2/2] Устанавливаю зависимости...
    .venv\Scripts\pip install -r requirements.txt -q
)

echo.
echo === Запуск бота ===
echo Excel: см. config.yaml - paths.excel_file
echo.

.venv\Scripts\python.exe -m src.main
set EXIT_CODE=%ERRORLEVEL%

echo.
if %EXIT_CODE%==0 (
    echo Готово. Файлы: output\docx, output\pdf, output\reports
) else (
    echo Завершено с ошибками. Смотрите logs\app.log
)
echo.
pause
exit /b %EXIT_CODE%
