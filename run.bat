@echo off
title Maha AI Olympiad - Register Scanner
cd /d "%~dp0"

echo ============================================================
echo   Maha AI Olympiad - Attendance Register Scanner
echo ============================================================
echo.

where node >nul 2>&1
if %errorlevel% equ 0 (
    echo Starting with Node.js...
    node server.js
    goto :eof
)

where py >nul 2>&1
if %errorlevel% equ 0 (
    echo Starting with Python Launcher...
    py -3.13 server.py
    goto :eof
)

echo Opening index.html in browser...
start "" "index.html"
pause
