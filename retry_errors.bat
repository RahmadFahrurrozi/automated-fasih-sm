@echo off
title BOT FASIH - RETRY DATA ERROR / BELUM SELESAI
echo ===================================================
echo     BOT FASIH: RETRY KHUSUS DATA GAGAL
echo ===================================================
echo.
echo Menjalankan bot dalam mode retry...
call .venv\Scripts\activate
python bot_fasih_perbaikan.py --retry-failed %*
echo.
pause
