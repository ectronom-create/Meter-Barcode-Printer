@echo off
title Meter Barcode Printer - Zebra ZT411
echo ========================================================
echo   Starting Meter Barcode Printer (Zebra ZT411 Direct)
echo ========================================================
start /b python server.py
timeout /t 1 /nobreak >nul
start "" "http://localhost:5000"
exit
