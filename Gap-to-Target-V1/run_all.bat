@echo off
setlocal EnableExtensions

cd /d "%~dp0"
set "PYTHONUTF8=1"
set "MPLBACKEND=Agg"

echo ============================================================
echo Gap-to-Target replication pipeline
echo ============================================================

if not exist ".venv\Scripts\python.exe" (
    echo [SETUP] Creating isolated Python environment...
    where py >nul 2>&1
    if not errorlevel 1 (
        py -3 -m venv .venv
    ) else (
        where python >nul 2>&1
        if errorlevel 1 (
            echo [ERROR] Python 3 was not found. Install Python and try again.
            exit /b 1
        )
        python -m venv .venv
    )
    if errorlevel 1 (
        echo [ERROR] The Python environment could not be created.
        exit /b 1
    )
)

echo [SETUP] Checking Python dependencies...
".venv\Scripts\python.exe" -m pip install --disable-pip-version-check --quiet -r requirements.lock
if errorlevel 1 (
    echo [ERROR] Required Python packages could not be installed.
    exit /b 1
)

".venv\Scripts\python.exe" code\master.py
if errorlevel 1 (
    echo.
    echo [ERROR] Replication stopped. See logs\replication.log for details.
    exit /b 1
)

echo.
echo ============================================================
echo Replication completed successfully.
echo Final materials are available in the outputs directory.
echo ============================================================
exit /b 0
