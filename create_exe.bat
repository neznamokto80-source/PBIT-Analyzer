@echo off
chcp 65001 >nul

:: === НАСТРОЙКИ ===
:: Имя исходного скрипта (без расширения, для поиска в _work)
set "APP_NAME=last_work_gui_work"

:: Имя выходного EXE файла
:: Если оставить пустым, будет использовано значение APP_NAME
set "OUTPUT_NAME=pbi_analyzer_v16"

:: =================================

:: Если OUTPUT_NAME не задан, используем APP_NAME
if "%OUTPUT_NAME%"=="" (
    set "OUTPUT_NAME=%APP_NAME%"
    echo [ℹ️] Имя выходного файла не задано. Используется: %OUTPUT_NAME%
) else (
    echo [ℹ️] Имя выходного файла: %OUTPUT_NAME%
)

title Компиляция: %APP_NAME% -> %OUTPUT_NAME%.exe

echo ========================================
echo    Исходник:  %APP_NAME%.py
echo    Выходной:  %OUTPUT_NAME%.exe
echo ========================================
echo.

:: 1. Проверка наличия PyInstaller
where pyinstaller >nul 2>&1
if %errorlevel% neq 0 (
    echo [⚠️] PyInstaller не найден. Устанавливаю...
    pip install pyinstaller
) else (
    echo [✅] PyInstaller найден.
)
echo.

:: 2. Установка зависимостей из requirements.txt
if exist requirements.txt (
    echo [📦] Установка зависимостей из requirements.txt...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo [❌] Ошибка при установке зависимостей!
        pause
        exit /b 1
    )
    echo [✅] Зависимости установлены.
) else (
    echo [⚠️] Файл requirements.txt не найден.
)
echo.

:: 3. Очистка старых папок сборки
echo [🧹] Очистка предыдущих сборок...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "%OUTPUT_NAME%.spec" del /q "%OUTPUT_NAME%.spec"
echo [✅] Готово
echo.

:: 4. Проверка существования исходного файла
if not exist "_work\%APP_NAME%.py" (
    echo [❌] Исходный файл _work\%APP_NAME%.py не найден!
    echo Текущая директория: %CD%
    pause
    exit /b 1
)

:: 5. Компиляция
echo [🚀] Компиляция в EXE...
pyinstaller --onefile ^
            --name "%OUTPUT_NAME%" ^
            --console ^
            --hidden-import PyQt6 ^
            "_work\%APP_NAME%.py"

if %errorlevel% neq 0 (
    echo [❌] Ошибка компиляции PyInstaller!
    pause
    exit /b 1
)

:: 6. Проверка результата
echo [🔍] Проверка результата...
if exist "dist\%OUTPUT_NAME%.exe" (
    echo [✅] Компиляция успешна!
    
    :: Получение размера файла
    for %%A in ("dist\%OUTPUT_NAME%.exe") do set size=%%~zA
    set /a sizeMB=%size% / 1048576
    echo Размер файла: %sizeMB% МБ
    
    :: Копирование в текущую папку
    echo [📂] Копирование %OUTPUT_NAME%.exe в текущую папку...
    copy "dist\%OUTPUT_NAME%.exe" . >nul
    if %errorlevel% equ 0 (
        echo [✅] Скопировано.
    ) else (
        echo [⚠️] Ошибка копирования (файл может быть занят).
    )
    
    :: Очистка временных файлов
    echo [🗑️] Очистка временных файлов...
    if exist build rmdir /s /q build
    if exist dist rmdir /s /q dist
    if exist "%OUTPUT_NAME%.spec" del /q "%OUTPUT_NAME%.spec"
    echo [✅] Готово
    
    echo.
    echo ========================================
    echo    ГОТОВО!
    echo    Файл: %OUTPUT_NAME%.exe
    echo    Путь: %CD%\%OUTPUT_NAME%.exe
    echo ========================================
) else (
    echo [❌] Ошибка: Файл dist\%OUTPUT_NAME%.exe не найден после компиляции.
    echo Проверьте логи выше.
)

pause
