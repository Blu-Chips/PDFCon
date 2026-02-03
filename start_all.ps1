# PDFCon Development Environment Startup Script
Write-Host "========================================" -ForegroundColor Cyan
Write-Host "Starting PDFCon Development Environment" -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan

# Set environment variables
$env:DATABASE_URL = "postgresql+asyncpg://postgres:postgres@localhost:5433/pdfcon"

# Start Backend
Write-Host "Starting Backend Server..." -ForegroundColor Green
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PWD\backend'; & '../.venv/Scripts/Activate.ps1'; python -m uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload" -WindowStyle Normal

Start-Sleep -Seconds 3

# Start Frontend
Write-Host "Starting Frontend Server..." -ForegroundColor Green
Start-Process powershell -ArgumentList "-NoExit", "-Command", "cd '$PWD\frontend'; npm run dev" -WindowStyle Normal

Write-Host "========================================" -ForegroundColor Cyan
Write-Host "PDFCon is now running!" -ForegroundColor Yellow
Write-Host "Backend: http://localhost:8000" -ForegroundColor White
Write-Host "Frontend: http://localhost:5173" -ForegroundColor White
Write-Host "API Docs: http://localhost:8000/api/docs" -ForegroundColor White
Write-Host "========================================" -ForegroundColor Cyan

Write-Host "`nPress Enter to stop all servers..." -ForegroundColor Red
Read-Host

# Stop all Python and Node processes
Write-Host "Stopping servers..." -ForegroundColor Yellow
Stop-Process -Name "python" -Force -ErrorAction SilentlyContinue
Stop-Process -Name "node" -Force -ErrorAction SilentlyContinue