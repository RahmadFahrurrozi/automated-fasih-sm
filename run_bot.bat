@echo off
REM Menjalankan bot otomasi menggunakan virtual environment .venv
set VENV_PYTHON=%~dp0.venv\Scripts\python.exe

if not exist "%VENV_PYTHON%" (
    echo [ERROR] Virtual environment .venv tidak ditemukan!
    echo Silakan jalankan: python -m venv .venv dan install dependencies.
    pause
    exit /b 1
)

echo Menjalankan bot_fasih_perbaikan.py dengan virtual environment...
echo Parameter: %*
"%VENV_PYTHON%" "%~dp0bot_fasih_perbaikan.py" %*

pause
