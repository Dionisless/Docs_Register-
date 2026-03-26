@echo off
chcp 65001 > nul
echo ============================================================
echo  Сборка LanDocs_Registrator.exe
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
echo [1/5] Установка зависимостей...
pip install pywin32 openpyxl pyinstaller python-docx tkinterdnd2 lxml selenium pyautogui
if errorlevel 1 (
    echo ОШИБКА: Не удалось установить зависимости.
    pause
    exit /b 1
)

:: Запускаем post-install для pywin32
python -m pywin32_postinstall -install >nul 2>&1

:: Определяем путь к tkinterdnd2
echo.
echo [2/5] Поиск пакета tkinterdnd2...
for /f "delims=" %%i in ('python -c "import tkinterdnd2, os; print(os.path.dirname(tkinterdnd2.__file__))"') do set TKDND_PATH=%%i
echo  Найден: %TKDND_PATH%

:: Собираем exe
echo.
echo [3/5] Сборка LanDocs_Registrator.exe через PyInstaller...
pyinstaller ^
    --onefile ^
    --windowed ^
    --name "LanDocs_Registrator" ^
    --add-data "%TKDND_PATH%;tkinterdnd2" ^
    --hidden-import "tkinterdnd2" ^
    --hidden-import "docx" ^
    --hidden-import "lxml" ^
    --hidden-import "selenium" ^
    --hidden-import "pyautogui" ^
    --hidden-import "win32com.client" ^
    --hidden-import "win32clipboard" ^
    landocs_register.py

if errorlevel 1 (
    echo ОШИБКА: Сборка не удалась. Смотрите лог выше.
    pause
    exit /b 1
)

:: Копируем yandexdriver.exe рядом с exe (если он есть)
echo.
echo [4/5] Проверка yandexdriver.exe...
if exist "yandexdriver.exe" (
    copy /y "yandexdriver.exe" "dist\yandexdriver.exe" >nul
    echo  Скопирован yandexdriver.exe в dist\
) else (
    echo  ВНИМАНИЕ: yandexdriver.exe не найден. Нужен только для вкладки "7 ДЭБ".
)

echo.
echo [5/5] Готово!
echo.
echo  Исполняемый файл: dist\LanDocs_Registrator.exe
echo.
echo  Для сетевого диска: скопируйте dist\LanDocs_Registrator.exe на сетевой диск.
echo  settings.json будет создан автоматически рядом с exe при первом сохранении настроек.
echo.
echo  ВАЖНО: для работы вкладки "7 ДЭБ" необходимо:
echo   - yandexdriver.exe рядом с exe
echo   - Яндекс Браузер установлен на целевом ПК
echo.
pause
