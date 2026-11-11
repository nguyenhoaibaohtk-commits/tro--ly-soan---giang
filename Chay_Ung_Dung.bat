@echo off
cd /d "%~dp0"
title EduKit 2025 - Tro Ly Giao Vien THPT
echo ===================================================
echo          EDUKIT 2025 - TRO LY GIAO VIEN
echo ===================================================
echo Dang khoi dong ung dung...
echo Trinh duyet web se tu dong mo trong giay lat.
echo Vui long KHONG tat cua so nay trong khi dung app.
echo.
"C:\Users\PL LAPTOP\AppData\Local\Programs\Python\Python312\python.exe" app.py
if %errorlevel% neq 0 (
    python app.py
)
pause
