@echo off
title SETUP BOT - INSTALASI DI LAPTOP BARU
cd /d "%~dp0"
echo ===================================================================
echo             INSTALASI DEPENDENSI BOT PERTAMINA
echo ===================================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python belum terinstall atau belum dicentang 'Add python.exe to PATH'!
    echo.
    echo Silakan download dan install Python di: https://www.python.org/downloads/
    echo PENTING: Saat install, centang pilihan "Add python.exe to PATH".
    echo.
    pause
    exit /b
)

echo [1/3] Python terdeteksi. Memperbarui pip dan menginstall library...
python -m pip install --upgrade pip
python -m pip install playwright openpyxl

echo.
echo [2/3] Mengunduh browser otomatis (Chromium)...
python -m playwright install chromium

echo.
echo ===================================================================
echo [3/3] SUKSES! Semua kebutuhan bot sudah terpasang.
echo Anda sekarang bisa langsung menjalankan bot lewat: JALANKAN_BOT.bat
echo ===================================================================
echo.
pause
