#Requires -Version 5.0
# Dev Proxy Quick Start Script (PowerShell)
# Stops any existing proxy, then runs the Outlook Shredder Proxy for dev/testing

Set-StrictMode -Version Latest

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Outlook Shredder Proxy - Development Build" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Stop any existing proxy processes
Write-Host "Checking for existing proxy processes..." -ForegroundColor Yellow
$existingProxies = Get-Process -Name "OutlookShredder.Proxy" -ErrorAction SilentlyContinue
if ($existingProxies) {
    Write-Host "Found $($existingProxies.Count) existing proxy process(es). Stopping..." -ForegroundColor Yellow
    foreach ($proc in $existingProxies) {
        Write-Host "  Stopping PID $($proc.Id)..." -ForegroundColor Yellow
        Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue
        Start-Sleep -Milliseconds 500
    }
    Write-Host "Existing proxy process(es) stopped." -ForegroundColor Green
} else {
    Write-Host "No existing proxy processes found." -ForegroundColor Green
}

# Navigate to release build
$buildPath = "C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0"
$exePath = Join-Path $buildPath "OutlookShredder.Proxy.exe"

if (-not (Test-Path $exePath)) {
    Write-Host "ERROR: Build not found at $buildPath" -ForegroundColor Red
    Write-Host "Please run: dotnet build -c Release" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Build Location: $buildPath" -ForegroundColor Green
Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Configuration" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Default Provider:  Claude (claude-sonnet-4-6)"
Write-Host "Fallback Provider: Google (gemini-1.5-pro)"
Write-Host "Listen Address:    http://localhost:7000"
Write-Host "Logs:              $(Join-Path $buildPath '..\..\..\..\Logs\')"
Write-Host ""

Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Starting Proxy Service..." -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Change to build directory so relative paths work
Push-Location $buildPath
try {
    # Run the proxy
    & $exePath
} finally {
    Pop-Location
}

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Proxy Stopped" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Read-Host "Press Enter to exit"
