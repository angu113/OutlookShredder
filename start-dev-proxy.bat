@echo off
REM Dev Proxy Quick Start Script
REM Stop any existing proxy, then run the Outlook Shredder Proxy for local development/testing

setlocal enabledelayedexpansion

echo.
echo ===========================================================
echo   Outlook Shredder Proxy - Development Build
echo ===========================================================
echo.

REM Stop any existing proxy processes
echo Checking for existing proxy processes...
tasklist /FI "IMAGENAME eq OutlookShredder.Proxy.exe" 2>nul | find /I /N "OutlookShredder.Proxy.exe">nul
if "%ERRORLEVEL%"=="0" (
    echo Found existing proxy process. Stopping...
    taskkill /IM OutlookShredder.Proxy.exe /F >nul 2>&1
    timeout /t 2 /nobreak >nul
    echo Existing proxy process stopped.
) else (
    echo No existing proxy processes found.
)

REM Navigate to release build
cd /d "C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0"

if not exist "OutlookShredder.Proxy.exe" (
    echo ERROR: Build not found at %cd%
    echo Please run: dotnet build -c Release
    pause
    exit /b 1
)

echo Build Location: %cd%
echo.
echo ===========================================================
echo   Configuration
echo ===========================================================
echo.
echo Default Provider:  Claude (claude-sonnet-4-6)
echo Fallback Provider: Google (gemini-1.5-pro)
echo Listen Address:    http://localhost:7000
echo Logs:              ..\..\..\Logs\
echo.

echo ===========================================================
echo   Starting Proxy Service...
echo ===========================================================
echo.

REM Run the proxy
OutlookShredder.Proxy.exe

echo.
echo ===========================================================
echo   Proxy Stopped
echo ===========================================================
pause
