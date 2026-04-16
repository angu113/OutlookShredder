#Requires -Version 5.0
# Dev Proxy Quick Start Script (PowerShell)
# Run the Outlook Shredder Proxy for local development/testing

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Outlook Shredder Proxy - Development Build" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host ""

# Navigate to release build
$buildPath = "C:\Users\angus\Shredder\Proxy\OutlookShredder\OutlookShredder.Proxy\bin\Release\net8.0"
$exePath = Join-Path $buildPath "OutlookShredder.Proxy.exe"

if (-not (Test-Path $exePath)) {
    Write-Host "ERROR: Build not found at $buildPath" -ForegroundColor Red
    Write-Host "Please run: dotnet build -c Release" -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

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

# Run the proxy
& $exePath

Write-Host ""
Write-Host "==========================================================" -ForegroundColor Cyan
Write-Host "  Proxy Stopped" -ForegroundColor Cyan
Write-Host "==========================================================" -ForegroundColor Cyan
Read-Host "Press Enter to exit"
