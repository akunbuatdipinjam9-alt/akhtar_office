@echo off
setlocal EnableExtensions
cd /d "%~dp0"
title Akhtar Office - Build Windows

color 0B
cls
echo ========================================================
echo           AKHTAR OFFICE - BUILD WINDOWS .EXE
echo ========================================================
echo.
echo Folder project:
echo %CD%
echo.

where node >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Node.js belum terpasang atau belum masuk PATH.
  echo.
  echo Install Node.js LTS dulu, lalu buka file ini lagi.
  goto :fail
)

where npm >nul 2>&1
if errorlevel 1 (
  echo [ERROR] npm tidak ditemukan.
  goto :fail
)

if not exist "package.json" (
  echo [ERROR] package.json tidak ditemukan.
  echo Taruh BUILD-WINDOWS.cmd di folder utama Akhtar Office.
  goto :fail
)

echo [1/4] Cek dependency...
if exist "node_modules\electron\dist\electron.exe" if exist "node_modules\electron-builder" goto :deps_ready

echo [INFO] Dependency belum lengkap. Menjalankan npm install...
call npm install
if errorlevel 1 (
  echo.
  echo [WARN] npm install gagal.
  echo Kalau error download Electron / ECONNRESET, kita bisa coba mirror.
  choice /C YN /N /M "Coba ulang lewat mirror npmmirror.com? [Y/N]: "
  if errorlevel 2 goto :fail
  set "ELECTRON_MIRROR=https://npmmirror.com/mirrors/electron/"
  call npm install
  set "ELECTRON_MIRROR="
  if errorlevel 1 goto :fail
)

:deps_ready
echo [OK] Dependency siap.
echo.

echo [2/4] Cek Electron...
if not exist "node_modules\electron\dist\electron.exe" (
  echo [INFO] Binary Electron belum ada. Mencoba npm rebuild electron...
  call npm rebuild electron
  if errorlevel 1 goto :fail
)

echo [OK] Electron siap.
echo.

echo [3/4] Tes syntax project...
node --check main.js
if errorlevel 1 goto :fail
node --check preload.js
if errorlevel 1 goto :fail
echo [OK] main.js dan preload.js lolos pengecekan.
echo.

echo [4/4] Membuat installer Windows...
call npm run dist:win
if errorlevel 1 (
  echo.
  echo [ERROR] Build gagal. Lihat pesan error di atas.
  goto :fail
)

echo.
echo ========================================================
echo                     BUILD BERHASIL
echo ========================================================
echo.
echo Installer ada di:
echo   %CD%\release\
echo.
if exist "release" start "" explorer "%CD%\release"
echo Tinggal double-click Akhtar-Office-Setup-*.exe
pause
exit /b 0

:fail
echo.
echo ========================================================
echo                     BUILD GAGAL
echo ========================================================
echo Jangan jalankan npm audit fix --force.
echo Copy bagian error terakhir kalau mau diperiksa.
echo.
pause
exit /b 1
