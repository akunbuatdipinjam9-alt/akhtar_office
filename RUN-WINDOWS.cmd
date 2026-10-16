@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Akhtar Office

where node >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Node.js belum terpasang.
  echo Install Node.js LTS dulu.
  pause
  exit /b 1
)

if not exist "package.json" (
  echo [ERROR] package.json tidak ditemukan.
  pause
  exit /b 1
)

if not exist "node_modules\docx\package.json" (
  echo Pertama kali dijalankan - memasang dependency...
  call npm install
  if errorlevel 1 (
    echo.
    echo [ERROR] npm install gagal.
    pause
    exit /b 1
  )
)

start "Akhtar Office" cmd /c "npm start"
exit /b 0
