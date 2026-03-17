@echo off
:: ============================================================================
:: build_all.bat — сборка всех 4 программ в .exe через PyInstaller
:: ============================================================================
:: Требования: pip install pyinstaller
:: Запускать из папки проекта
:: ============================================================================

setlocal
set DIST=dist
set SPEC_OPTS=--onefile --windowed --clean

echo ============================================================
echo  Сборка программ Docs_Register
echo ============================================================

:: 1 — Регистратор писем
echo.
echo [1/4] 1_letter_register.py ...
pyinstaller %SPEC_OPTS% ^
    --name "1_LetterRegister" ^
    --add-data "shared_lib.py;." ^
    1_letter_register.py
if errorlevel 1 (
    echo ОШИБКА при сборке 1_letter_register
    goto :error
)

:: 2 — Раскладка по папкам
echo.
echo [2/4] 2_ustavki_folders.py ...
pyinstaller %SPEC_OPTS% ^
    --name "2_UstavkiFolders" ^
    --add-data "shared_lib.py;." ^
    2_ustavki_folders.py
if errorlevel 1 (
    echo ОШИБКА при сборке 2_ustavki_folders
    goto :error
)

:: 3 — Карты Visio
echo.
echo [3/4] 3_ustavki_map.py ...
pyinstaller %SPEC_OPTS% ^
    --name "3_UstavkiMap" ^
    --add-data "shared_lib.py;." ^
    3_ustavki_map.py
if errorlevel 1 (
    echo ОШИБКА при сборке 3_ustavki_map
    goto :error
)

:: 4 — ДЭБ Selenium
echo.
echo [4/4] 4_deb_selenium.py ...
pyinstaller %SPEC_OPTS% ^
    --name "4_DebSelenium" ^
    --add-data "shared_lib.py;." ^
    4_deb_selenium.py
if errorlevel 1 (
    echo ОШИБКА при сборке 4_deb_selenium
    goto :error
)

echo.
echo ============================================================
echo  Готово! .exe файлы в папке: %DIST%\
echo ============================================================
echo   dist\1_LetterRegister.exe
echo   dist\2_UstavkiFolders.exe
echo   dist\3_UstavkiMap.exe
echo   dist\4_DebSelenium.exe
echo ============================================================
goto :eof

:error
echo.
echo Сборка завершилась с ошибкой.
exit /b 1
