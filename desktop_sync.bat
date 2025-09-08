@echo off
REM Desktop Calendar Sync - Ad-hoc execution
title Outlook Calendar Sync

REM Change to the script directory
cd /d "%~dp0"

REM Run the desktop sync script
"%~dp0.venv\Scripts\python.exe" "%~dp0desktop_sync.py"
