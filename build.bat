@echo off
chcp 65001 > nul
echo ============================================================
echo  Сборка landocs_register.exe
echo ============================================================
echo.

:: Проверяем наличие Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ОШИБКА: Python не найден. Установите Python 3.8+ и добавьте в PATH.
    pause
    exit /b 1
)

:: Устанавливаем зависимости
echo [1/4] Установка зависимостей...
pip install pywin32 openpyxl pyinstaller python-docx tkinterdnd2 selenium pyautogui
if errorlevel 1 (
    echo ОШИБКА: Не удалось установить зависимости.
    pause
    exit /b 1
)

:: Запускаем post-install для pywin32 (регистрирует DLL)
python -m pywin32_postinstall -install >nul 2>&1

:: Определяем путь к tkinterdnd2 для включения в сборку
echo.
echo [2/4] Поиск пакета tkinterdnd2...
for /f "delims=" %%i in ('python -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))"') do set TKDND_PATH=%%i
echo  Найден: %TKDND_PATH%

:: Собираем .exe
echo.
echo [3/4] Сборка .exe через PyInstaller...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "LanDocs_Registrator" ^
    --icon NONE ^
    --add-data "%TKDND_PATH%;tkinterdnd2" ^
    --hidden-import "tkinterdnd2" ^
    --hidden-import "docx" ^
    --hidden-import "selenium" ^
    --hidden-import "pyautogui" ^
    landocs_register.py

if errorlevel 1 (
    echo ОШИБКА: Сборка не удалась. Смотрите лог выше.
    pause
    exit /b 1
)

echo.
echo [4/4] Готово!
echo  Исполняемый файл: dist\LanDocs_Registrator.exe
echo.
echo  Скопируйте dist\LanDocs_Registrator.exe в удобное место
echo  и настройте горячую клавишу через hotkey.ahk или ярлык.
echo.
pause
