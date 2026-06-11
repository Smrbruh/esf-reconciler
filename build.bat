@echo off
chcp 65001 > nul
echo.
echo ============================================================
echo  Сборка ЭСФ_Сверка.exe через PyInstaller
echo ============================================================
echo.

:: Проверяем наличие PyInstaller
pip show pyinstaller > nul 2>&1
if errorlevel 1 (
    echo [!] PyInstaller не установлен. Устанавливаю...
    pip install pyinstaller
    echo.
)

:: Сборка .exe
pyinstaller --onefile ^
            --name "ЭСФ_Сверка" ^
            --add-data "config.yaml;." ^
            esf_reconcile.py

if errorlevel 1 (
    echo.
    echo [ОШИБКА] Сборка завершилась с ошибкой.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  Готово! Файл: dist\ЭСФ_Сверка.exe
echo  Скопируйте его вместе с папками input\ и output\
echo  и файлом config.yaml в рабочую директорию.
echo ============================================================
echo.
pause
