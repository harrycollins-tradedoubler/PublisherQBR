# Start Publisher QBR legacy app-hub services

Write-Host "Starting Publisher QBR legacy app-hub services..." -ForegroundColor Cyan
Write-Host "Active QBR generation uses the Chrome extension, local runner on 3020, and qbr-pptx-service on 3010." -ForegroundColor Cyan

Write-Host "Starting backend server..." -ForegroundColor Yellow
$backendPath = Join-Path $PSScriptRoot "..\backend"
if (Test-Path $backendPath) {
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$backendPath'; .\.venv\Scripts\Activate.ps1; .\.venv\Scripts\python.exe -m uvicorn app.main:app --reload --host 0.0.0.0 --port 8000"
} else {
    Write-Host "Backend folder not found." -ForegroundColor Red
}

Write-Host "Starting frontend server..." -ForegroundColor Yellow
$frontendPath = Join-Path $PSScriptRoot "..\frontend"
if (Test-Path $frontendPath) {
    Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$frontendPath'; npm run dev"
} else {
    Write-Host "Frontend folder not found." -ForegroundColor Red
}

Write-Host ""
Write-Host "Legacy app-hub services starting..." -ForegroundColor Green
Write-Host "Backend: http://localhost:8000" -ForegroundColor Cyan
Write-Host "Frontend: http://localhost:5173" -ForegroundColor Cyan
