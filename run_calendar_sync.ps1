# PowerShell script for calendar sync automation
Set-Location "C:\OutlookCalendarExports"

Write-Host "Starting Outlook Calendar Export at $(Get-Date)" -ForegroundColor Green

try {
    # Activate virtual environment
    & ".\.venv\Scripts\Activate.ps1"
    
    # Run the export process
    Write-Host "Step 1: Exporting from Outlook..." -ForegroundColor Yellow
    python export_outlook_calendar.py
    if ($LASTEXITCODE -ne 0) { throw "Failed to export from Outlook" }
    
    Write-Host "Step 2: Converting to ICS..." -ForegroundColor Yellow
    python csv_to_ics.py
    if ($LASTEXITCODE -ne 0) { throw "Failed to convert to ICS" }
    
    Write-Host "Step 3: Emailing ICS file..." -ForegroundColor Yellow
    python email_icloud.py
    if ($LASTEXITCODE -ne 0) { throw "Failed to send email" }
    
    Write-Host "Calendar export completed successfully at $(Get-Date)" -ForegroundColor Green
    
} catch {
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}
