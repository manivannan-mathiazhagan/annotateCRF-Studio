@echo off
cd /d "%~dp0"

echo =====================================
echo    AnnotateCRF Studio
echo =====================================

IF NOT EXIST annotatecrf_studio.py (
    echo ERROR: annotatecrf_studio.py not found!
    pause
    exit /b
)

python annotatecrf_studio.py

IF %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Application crashed.
)

pause