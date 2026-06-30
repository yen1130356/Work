@echo off
title 🏥 台中榮總傳送系統 - 伺服器整合啟動器

echo ⏳ 正在啟動免安裝版 PostgreSQL 資料庫伺服器...
start /b "" "C:\pgsql\bin\pg_ctl.exe" -D "C:\pgsql\data" start

echo ⏳ 預留 3 秒讓資料庫完成開機就緒...
timeout /t 3 /nobreak > null

echo ⏳ 正在喚醒傳送數據自動化大腦 (Flask)...
start http://localhost:5000

cd /d "%~dp0"
python system.py
pause