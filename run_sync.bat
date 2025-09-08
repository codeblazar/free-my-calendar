@echo off
REM Create log file with timestamp
set LOGFILE="%~dp0sync_log_%date:~-4,4%_%date:~-10,2%_%date:~-7,2%.txt"
echo Starting Outlook Calendar Sync with Deletion Tracking at %date% %time% > %LOGFILE%
echo Starting Outlook Calendar Sync with Deletion Tracking at %date% %time%

REM Change to the script directory (critical for Task Scheduler)
cd /d "%~dp0"
echo Current directory: %cd% >> %LOGFILE%
echo Current directory: %cd%

REM Check if virtual environment exists
if not exist ".venv\Scripts\activate.bat" (
    echo ERROR: Virtual environment not found at %~dp0.venv >> %LOGFILE%
    echo ERROR: Virtual environment not found at %~dp0.venv
    echo Current directory: %cd% >> %LOGFILE%
    echo Current directory: %cd%
    echo Please ensure the batch file is in the correct directory >> %LOGFILE%
    echo Please ensure the batch file is in the correct directory
    exit /b 1
)

REM Activate virtual environment using full path
echo Activating virtual environment... >> %LOGFILE%
echo Activating virtual environment...
call "%~dp0.venv\Scripts\activate.bat"

REM Verify Python is from virtual environment
echo Python location: >> %LOGFILE%
where python >> %LOGFILE%
echo Python location: 
where python

REM Run the sync process (handles export, conversion, deletion tracking, and email)
echo Running sync process... >> %LOGFILE%
echo Running sync process...
python sync.py >> %LOGFILE% 2>&1

if %ERRORLEVEL% EQU 0 (
    echo Calendar Sync completed successfully at %date% %time% >> %LOGFILE%
    echo Calendar Sync completed successfully at %date% %time%
) else (
    echo Calendar Sync failed at %date% %time% >> %LOGFILE%
    echo Calendar Sync failed at %date% %time%
    echo Error level: %ERRORLEVEL% >> %LOGFILE%
    echo Error level: %ERRORLEVEL%
    exit /b 1
)

REM Log completion
echo Sync process finished at %date% %time% >> %LOGFILE%
echo Sync process finished at %date% %time%
