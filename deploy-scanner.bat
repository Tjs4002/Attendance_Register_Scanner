@echo off
title Deploy Attendance Register Scanner
cd /d "%~dp0"

echo ============================================================
echo   Maha AI Olympiad - Deploy Register Scanner
echo ============================================================
echo.
echo Select your deployment platform:
echo   [1] Deploy to Vercel (Fastest, automated public URL)
echo   [2] Deploy to Netlify (Production deploy via CLI)
echo   [3] Open Netlify Drop in Browser (0 install, drag-and-drop)
echo   [4] Test locally first
echo.
set /p choice="Enter choice (1-4): "

if "%choice%"=="1" (
    echo.
    echo Deploying to Vercel with npx vercel...
    npx vercel
    pause
    goto :eof
)

if "%choice%"=="2" (
    echo.
    echo Deploying to Netlify with npx netlify...
    npx netlify deploy --dir=.
    pause
    goto :eof
)

if "%choice%"=="3" (
    echo.
    echo Opening Netlify Drop...
    echo Just drag the 'attendance-register-scanner' folder into the browser!
    start "" "https://app.netlify.com/drop"
    pause
    goto :eof
)

if "%choice%"=="4" (
    start "" "index.html"
    goto :eof
)

echo Invalid choice.
pause
