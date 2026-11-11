@echo off
cd /d "%~dp0"
title Day code len GitHub - EduKit 2025
echo ===================================================
echo     DANG DAY MA NGUON LEN GITHUB CUA THAY BAO
echo ===================================================
echo.
echo Kho luu tru: https://github.com/nguyenhoaibaohtk-commits/tro--ly-soan---giang.git
echo.
git remote remove origin 2>nul
git remote add origin https://github.com/nguyenhoaibaohtk-commits/tro--ly-soan---giang.git
git branch -M main
echo Dang ket noi va day code len...
git push -u origin main
echo.
if %errorlevel% equ 0 (
    echo ===================================================
    echo   [THANH CONG] Da day code len GitHub thanh cong!
    echo ===================================================
) else (
    echo ===================================================
    echo   Neu co cua so yeu cau dang nhap GitHub, 
    echo   ban vui long dang nhap de hoan tat nhe.
    echo ===================================================
)
echo.
pause
