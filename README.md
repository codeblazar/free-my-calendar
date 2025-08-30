# Outlook Calendar Export Tool

A Python tool to automatically export Microsoft Outlook calendar events, convert them to iCalendar format, and email them to your iCloud account with proper sync capabilities.

## Features

- 🔄 **Automated Export**: Extract calendar events from running Outlook application
- 📅 **iCalendar Format**: Convert to standard ICS format with unique event IDs
- 📧 **Email Delivery**: Automatically send to your iCloud account
- ⏰ **Scheduled Sync**: Run twice daily (12PM & 9PM) via Windows Task Scheduler
- 🔧 **Configurable**: All settings managed via environment variables
- 🆔 **Smart Sync**: Unique event IDs prevent duplicates and enable proper updates
- 📱 **Multi-Calendar**: Designed to work with separate work/personal calendars

## Requirements

- Windows 10/11
- Microsoft Outlook (must be running and signed in)
- Python 3.7+
- iCloud account with app-specific password

## Quick Start

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/outlook-calendar-export.git
   cd outlook-calendar-export
   ```

2. **Set up Python environment**
   ```powershell
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

3. **Configure environment variables**
   Copy `.env.example` to `.env` and update with your credentials:
   ```env
   ICLOUD_EMAIL=your_email@icloud.com
   ICLOUD_APP_PASSWORD=your_app_specific_password
   ```

4. **Run manual sync** (test)
   ```powershell
   .\run_calendar_sync.bat
   ```

5. **Set up automation** (optional)
   Follow the automation guide below to schedule automatic syncing.

## Configuration

All settings are configured via the `.env` file:

### Required Settings
```env
ICLOUD_EMAIL=your_email@icloud.com
ICLOUD_APP_PASSWORD=your_app_specific_password
```

### Optional Settings (with defaults)
```env
# Export settings
EXPORT_DAYS=30
EXPORT_DIRECTORY=C:\OutlookCalendarExports
CSV_FILENAME=outlook_calendar_export.csv
ICS_FILENAME=outlook_calendar_export.ics
BODY_CHAR_LIMIT=500

# Calendar metadata
CALENDAR_NAME=Pete Work
CALENDAR_DESCRIPTION=Corporate Outlook Calendar Export

# Email settings
EMAIL_SUBJECT=Automated Outlook Calendar Export
EMAIL_BODY=Find attached the latest Outlook calendar export as iCal.

# SMTP settings
SMTP_SERVER=smtp.mail.me.com
SMTP_PORT=465
SMTP_TIMEOUT=30
```

## Usage

### Manual Export (3-step process)

1. **Export from Outlook**
   ```powershell
   python export_outlook_calendar.py
   ```

2. **Convert to iCalendar**
   ```powershell
   python csv_to_ics.py
   ```

3. **Email the file**
   ```powershell
   python email_icloud.py
   ```

### Automated Export

Use the provided batch script:
```powershell
.\run_calendar_sync.bat
```

## Automation Setup

Set up Windows Task Scheduler for automatic twice-daily syncing:

1. Open Task Scheduler (`Win+R` → `taskschd.msc`)
2. Create two tasks:
   - **12PM Task**: Daily at 12:00 PM
   - **9PM Task**: Daily at 9:00 PM
3. Point both to: `C:\path\to\run_calendar_sync.bat`
4. Configure to run in background whether user is logged in or not

### Important Notes for Automation
- ✅ Microsoft Outlook must be running
- ✅ Computer must be powered on
- ❌ No other applications need to be open
- ⚠️ Missed runs are skipped (no catch-up)

## iCloud Setup

1. **Enable 2-Factor Authentication** on your Apple ID
2. **Generate App-Specific Password**:
   - Go to appleid.apple.com
   - Sign In → App-Specific Passwords
   - Generate password for "Outlook Calendar Export"
3. **Create Separate Calendar** in iCal:
   - Recommended: Create "Pete Work" or "Work Calendar"
   - Import ICS files to this dedicated calendar
4. **Import Strategy**:
   - Replace/overwrite the work calendar each time
   - This ensures deleted Outlook events are removed from iCal

## Project Structure

```
outlook-calendar-export/
├── .env.example              # Environment variables template
├── .gitignore               # Git ignore rules
├── README.md                # This file
├── requirements.txt         # Python dependencies
├── export_outlook_calendar.py   # Step 1: Export from Outlook
├── csv_to_ics.py           # Step 2: Convert to iCalendar
├── email_icloud.py         # Step 3: Email the file
├── run_calendar_sync.bat   # Automation script (Windows)
├── run_calendar_sync.ps1   # PowerShell alternative
└── calendar_scheduler.py   # Python scheduler (alternative)
```

## Sync Behavior

This tool generates **unique event IDs** for proper calendar synchronization:

- ✅ **New Events**: Added to calendar
- ✅ **Updated Events**: Modified in calendar (same UID)
- ✅ **Deleted Events**: Removed when calendar is replaced
- ✅ **No Duplicates**: Same event won't be imported multiple times

## Troubleshooting

### Common Issues

**Authentication Error**
- Verify iCloud email and app-specific password in `.env`
- Ensure 2FA is enabled on Apple ID

**Outlook Connection Failed**  
- Ensure Microsoft Outlook is running and signed in
- Check Windows permissions for COM access

**Email Not Sending**
- Test network connectivity to smtp.mail.me.com:465
- Verify app-specific password is correct
- Check firewall/antivirus blocking SMTP

**File Not Found Errors**
- Run scripts in correct order: export → convert → email
- Check file paths in `.env` configuration

### Debug Mode

Enable detailed logging by running individual scripts and checking output.

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- Uses `win32com.client` for Outlook integration
- Built for Windows Task Scheduler automation
- Designed for iCloud calendar synchronization
