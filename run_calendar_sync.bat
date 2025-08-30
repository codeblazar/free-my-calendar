@echo off
cd /d "C:\OutlookCalendarExports"

echo Starting Outlook Calendar Export at %date% %time%

REM Activate virtual environment
call .venv\Scripts\activate.bat

REM Run the export process
echo Step 1: Exporting from Outlook...
python export_outlook_calendar.py
if errorlevel 1 (
    echo ERROR: Failed to export from Outlook
    exit /b 1
)

echo Step 2: Converting to ICS...
python csv_to_ics.py
if errorlevel 1 (
    echo ERROR: Failed to convert to ICS
    exit /b 1
)

echo Step 3: Emailing ICS file...
python email_icloud.py
if errorlevel 1 (
    echo ERROR: Failed to send email
    exit /b 1
)

echo Calendar export completed successfully at %date% %time%
