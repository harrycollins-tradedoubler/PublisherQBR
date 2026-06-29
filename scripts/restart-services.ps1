# Restart Publisher QBR legacy app-hub services

Write-Host "Restarting Publisher QBR legacy app-hub services..." -ForegroundColor Cyan

$stopScript = Join-Path $PSScriptRoot "stop-services.ps1"
& $stopScript

Start-Sleep -Seconds 2

$startScript = Join-Path $PSScriptRoot "start-services.ps1"
& $startScript
