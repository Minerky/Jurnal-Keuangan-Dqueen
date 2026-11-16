@echo off
title Jurnal Keuangan DQueen Florist
cd /d "%~dp0"

echo ===================================================
echo   Memulai Jurnal Keuangan DQueen Florist...
echo   Membuka http://localhost:3000 di browser...
echo ===================================================

start "" http://localhost:3000
node server.js
pause
