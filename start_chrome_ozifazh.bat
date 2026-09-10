@echo off
echo ===================================================
echo   Menjalankan Chrome Debugging untuk Akun:
echo   ozifazh@gmail.com (C:\selenium\chrome_profile)
echo ===================================================
echo.

start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\selenium\chrome_profile"

echo.
echo Chrome berhasil dibuka dengan debugging port 9222!
echo Silakan:
echo 1. Nyalakan VPN BPS
echo 2. Buka dan login ke Fasih (fasih-sm.bps.go.id)
echo 3. Jalankan script python: python bot_fasih_perbaikan.py
echo.
pause
