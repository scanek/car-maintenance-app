@echo off
chcp 65001 >nul
echo ========================================================
echo   Запуск панели обслуживания Changan CS55 Plus
echo ========================================================
echo.

where docker >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    echo [1/2] Запуск через Docker Compose...
    docker compose up --build -d
    echo [2/2] Приложение запущено в Docker!
    timeout /t 2 >nul
    start http://localhost:8080
) else (
    echo [!] Docker не найден в PATH. Запуск в локальном режиме Python...
    echo [1/2] Установка / проверка зависимостей...
    python -m pip install -r requirements.txt >nul 2>nul
    echo [2/2] Запуск сервера на http://localhost:8080...
    start http://localhost:8080
    python app.py
)

pause
