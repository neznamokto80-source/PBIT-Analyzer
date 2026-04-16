@echo off
setlocal

:menu
cls
echo ==========================================
echo            PBIT Analyzer Tasks
echo ==========================================
echo.
echo 1. Run GUI application
echo    Starts: python _work\last_work_gui_work.py
echo.
echo 2. Lint code
echo    Runs:   python -m ruff check .
echo.
echo 3. Run tests
echo    Runs:   python -m pytest
echo.
echo 4. Check all (lint + test)
echo    Runs lint first, then tests
echo.
echo 5. Exit
echo.
set /p CHOICE=Enter option number [1-5]: 

if "%CHOICE%"=="1" goto run_app
if "%CHOICE%"=="2" goto lint
if "%CHOICE%"=="3" goto test
if "%CHOICE%"=="4" goto check_all
if "%CHOICE%"=="5" goto end

echo.
echo Invalid option. Please select 1, 2, 3, 4, or 5.
pause
goto menu

:run_app
echo.
echo [RUN] Starting GUI application...
python _work\last_work_gui_work.py
echo.
pause
goto menu

:lint
echo.
echo [LINT] Running Ruff...
python -m ruff check .
echo.
pause
goto menu

:test
echo.
echo [TEST] Running Pytest...
python -m pytest
echo.
pause
goto menu

:check_all
echo.
echo [CHECK] Running Ruff...
python -m ruff check .
if errorlevel 1 (
    echo.
    echo [CHECK] Lint failed. Tests skipped.
    pause
    goto menu
)
echo.
echo [CHECK] Running Pytest...
python -m pytest
echo.
pause
goto menu

:end
echo Exiting.
endlocal
exit /b 0
