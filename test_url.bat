@echo off
REM Menjalankan test khusus untuk URL assignment tertentu
set VENV_PYTHON=%~dp0.venv\Scripts\python.exe

set TARGET_URL=https://fasih-sm.bps.go.id/app/assignment/fd68e454-ba45-4b85-8205-f3bf777ded24/001c4782-dc60-4aee-94db-28f7a295da3e

echo Menjalankan testing untuk URL:
echo %TARGET_URL%
echo.

"%VENV_PYTHON%" "%~dp0bot_fasih_perbaikan.py" --url "%TARGET_URL%"

pause
