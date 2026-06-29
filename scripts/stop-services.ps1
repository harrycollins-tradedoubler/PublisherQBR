# Stop Publisher QBR legacy app-hub services

Write-Host "Stopping Publisher QBR legacy app-hub services..." -ForegroundColor Cyan

Write-Host "Stopping backend processes..." -ForegroundColor Yellow
Get-Process -Name "python" -ErrorAction SilentlyContinue | Stop-Process -Force
Get-Process -Name "uvicorn" -ErrorAction SilentlyContinue | Stop-Process -Force

Write-Host "Stopping frontend processes..." -ForegroundColor Yellow
$nodeProcesses = Get-NetTCPConnection -LocalPort 5173 -ErrorAction SilentlyContinue | Select-Object -ExpandProperty OwningProcess -Unique
foreach ($processId in $nodeProcesses) {
    Stop-Process -Id $processId -Force -ErrorAction SilentlyContinue
}

Get-Process -Name "node" -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowTitle -match "vite" -or $_.Path -match "node" } | Stop-Process -Force -ErrorAction SilentlyContinue

Write-Host "Legacy app-hub services stopped." -ForegroundColor Green
