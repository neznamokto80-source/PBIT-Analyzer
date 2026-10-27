@echo off
chcp 65001 >nul
:: Wrapper around build_exe.py: all options are passed through.
::   create_exe.bat             -> GUI + CLI (onefile)
::   create_exe.bat --gui       -> GUI only
::   create_exe.bat --cli       -> CLI only
::   create_exe.bat --dir       -> onedir build
::   create_exe.bat --no-copy   -> keep EXE in dist/ only
::   create_exe.bat --clean     -> remove build/, dist/ and *.spec
setlocal

set "PYTHON=py -3"
%PYTHON% --version >nul 2>&1
if errorlevel 1 (
    set "PYTHON=python"
    %PYTHON% --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python 3 not found. Install Python 3.10+ and try again.
        pause
        exit /b 1
    )
)

echo [INFO] Python: %PYTHON%
%PYTHON% build_exe.py %*
set "RESULT=%ERRORLEVEL%"

if not "%RESULT%"=="0" (
    echo [ERROR] Build failed with exit code %RESULT%.
    pause
    exit /b %RESULT%
)

echo [OK] Build finished. EXE is in the project root; dist\ stays only with --no-copy or --dir.
pause
exit /b 0
