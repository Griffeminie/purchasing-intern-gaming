@echo off
title Purchasing Tool Server
cd /d "%~dp0"

echo ============================================
echo   Starting Purchasing Tool Server...
echo ============================================
echo.
echo WARNING: CLOSING THIS WINDOW WILL STOP THE SERVER.
REM Check Node.js is installed
where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js is not installed or not in PATH.
    echo Please install Node.js from https://nodejs.org and try again.
    pause
    exit /b 1
)

REM Install dependencies if node_modules is missing (first run only)
if not exist "node_modules" (
    echo First-time setup: installing dependencies...
    call npm install
    echo.
)

REM ── Find this machine's LAN IPv4 address (skips localhost/loopback) ────────
set "IP="
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr /R /C:"IPv4 Address"') do (
    if not defined IP (
        set "candidate=%%a"
        setlocal enabledelayedexpansion
        set "candidate=!candidate: =!"
        if not "!candidate!"=="127.0.0.1" (
            endlocal
            set "IP=%%a"
        ) else (
            endlocal
        )
    )
)
set "IP=%IP: =%"
if not defined IP set "IP=localhost"

echo Server will be reachable at: http://%IP%:3000
echo.

REM Open the browser to the NETWORK address after a short delay, with no
REM visible extra window (PowerShell runs hidden instead of via a new cmd)
start "" /min powershell -WindowStyle Hidden -Command "Start-Sleep -Seconds 3; Start-Process 'http://%IP%:3000'"

REM Start the actual server (this keeps running — closing this window stops the server)
node server.js

echo.
echo Server stopped. Press any key to close this window.
pause >nul