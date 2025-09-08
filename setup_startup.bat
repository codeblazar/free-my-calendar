@echo off
echo 🚀 Setting up Calendar Sync to run at Windows startup
echo ====================================================

REM Get the startup folder path
set "StartupFolder=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

echo Startup folder: %StartupFolder%
echo.

REM Create a batch file that will run at startup
echo Creating startup script...

(
echo @echo off
echo REM Calendar Sync - Startup Runner
echo timeout /t 60 /nobreak ^>nul
echo echo Running Calendar Sync at startup...
echo cd /d "C:\OutlookCalendarExports"
echo call run_sync.bat
) > "%StartupFolder%\Calendar-Sync-Startup.bat"

if exist "%StartupFolder%\Calendar-Sync-Startup.bat" (
    echo ✅ Startup script created successfully!
    echo 📍 Location: %StartupFolder%\Calendar-Sync-Startup.bat
    echo.
    echo ⚠️  This will run calendar sync 60 seconds after Windows login
    echo 💡 To remove: Delete the file from the startup folder
) else (
    echo ❌ Failed to create startup script
)

echo.
pause
