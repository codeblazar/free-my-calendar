# FreeMyCalendar - Outlook to iCloud Sync

A comprehensive Python solution for syncing Microsoft Outlook calendar events to iCloud with full deletion tracking and smart sync capabilities.

## Features

- 🔄 **Smart Sync**: Full bidirectional sync with deletion tracking
- 📅 **iCalendar Format**: Standard ICS format with unique event IDs
- 📧 **Auto-Email**: Direct delivery to iCloud via SMTP
- 🎯 **Targeted Range**: Configurable date range (default: 2 weeks past to 12 weeks future)
- 🔧 **Outlook Management**: Auto-detects and starts Classic Outlook if needed
- 🖥️ **Desktop Integration**: One-click sync via desktop shortcut
- 🆔 **Duplicate Prevention**: MD5-based unique event IDs prevent duplicates
- 📱 **Cross-Platform**: Perfect for work/personal calendar separation

## Quick Start

1. **Clone and Setup**
   ```powershell
   git clone https://github.com/codeblazar/FreeMyCalendar.git
   cd FreeMyCalendar
   python -m venv .venv
   .\.venv\Scripts\Activate.ps1
   pip install -r requirements.txt
   ```

2. **Configure**
   ```powershell
   copy .env.example .env
   # Edit .env with your iCloud credentials
   ```

3. **Test Run**
   ```powershell
   .\desktop_sync.bat
   ```

## Configuration

Copy `.env.example` to `.env` and configure:

```env
# Required: iCloud credentials
ICLOUD_EMAIL=your_email@icloud.com
ICLOUD_APP_PASSWORD=your_app_specific_password

# Optional: Outlook settings
OUTLOOK_EMAIL=your_work@company.com
EXPORT_DAYS=98  # 2 weeks past + 12 weeks future

# Optional: Calendar metadata
CALENDAR_NAME=Work Calendar
CALENDAR_DESCRIPTION=Outlook Calendar Export
```

## How It Works

### Core Components

1. **sync.py** - Main orchestrator that coordinates the entire sync process
2. **outlook_manager.py** - Detects and manages Classic Outlook application
3. **export_outlook_calendar.py** - Extracts calendar events via COM interface
4. **sync_tracker.py** - Tracks deletions and modifications between syncs
5. **csv_to_ics.py** - Converts CSV export to iCalendar format with unique IDs
6. **email_icloud.py** - Sends ICS file to iCloud via SMTP

### Sync Process

1. **Outlook Check**: Ensures Classic Outlook is running
2. **Export**: Extracts events from specified date range
3. **Deletion Tracking**: Identifies removed events from previous sync
4. **Conversion**: Creates ICS file with unique event IDs
5. **Email**: Sends both main calendar and deletions to iCloud

## Usage Options

### Option 1: Desktop Shortcut (Recommended)
Run the setup to create a desktop shortcut:
```powershell
python create_shortcut.py
```
Then simply double-click "Sync Calendar.lnk" on your desktop.

### Option 2: Command Line
```powershell
# Full sync with deletion tracking
python sync.py

# Individual steps
python export_outlook_calendar.py
python csv_to_ics.py  
python email_icloud.py
```

### Option 3: Startup Automation
```powershell
.\setup_startup.bat
```
This creates a startup script (manual execution still required).

## Requirements

- **Windows 10/11**
- **Microsoft Outlook** (Classic version recommended)
- **Python 3.7+**
- **iCloud account** with app-specific password enabled

## iCloud Setup Guide

### 1. Enable App-Specific Passwords
1. Go to [appleid.apple.com](https://appleid.apple.com)
2. Sign in and navigate to "App-Specific Passwords"
3. Generate a new password labeled "Calendar Sync"
4. Copy this password to your `.env` file

### 2. Create Dedicated Calendar
1. Open Calendar app on Mac/iPhone/iCloud.com
2. Create new calendar: "Work Calendar" or similar
3. This will be your dedicated sync target

### 3. Import Strategy
- **Replace Method**: Overwrite the work calendar each sync
- **Benefits**: Automatically removes deleted Outlook events
- **Setup**: Import ICS files to your dedicated calendar only

## Project Structure

```
FreeMyCalendar/
├── 📁 Core Components
│   ├── sync.py                    # Main sync orchestrator
│   ├── outlook_manager.py         # Outlook process management  
│   ├── export_outlook_calendar.py # COM interface for Outlook
│   ├── sync_tracker.py           # Deletion tracking system
│   ├── csv_to_ics.py             # CSV to iCalendar converter
│   └── email_icloud.py           # SMTP email automation
├── 📁 User Interface
│   ├── desktop_sync.py           # Interactive sync with prompts
│   ├── desktop_sync.bat          # Batch wrapper
│   └── create_shortcut.py        # Desktop shortcut creator
├── 📁 Automation
│   ├── run_sync.bat              # Basic sync runner
│   └── setup_startup.bat         # Startup folder integration
├── 📁 Configuration
│   ├── .env.example              # Environment template
│   ├── requirements.txt          # Python dependencies
│   └── .gitignore               # Git ignore rules
└── 📁 Documentation
    ├── README.md                 # This file
    ├── CHANGELOG.md             # Version history
    └── LICENSE                  # MIT license
```

## Advanced Features

### Deletion Tracking
The sync system maintains a history of exported events and automatically detects:
- ✅ **New Events**: Added to calendar
- ✅ **Modified Events**: Updated with same unique ID  
- ✅ **Deleted Events**: Tracked and removed via separate ICS
- ✅ **No Duplicates**: MD5-based unique IDs prevent duplicates

### Date Range Optimization
By default, syncs:
- **Past**: 2 weeks (to catch late updates)
- **Future**: 12 weeks (to cover quarterly planning)
- **Configurable**: Adjust via `EXPORT_DAYS` setting

### Outlook Management
- **Auto-Detection**: Finds running Outlook instances
- **Process Management**: Starts Classic Outlook if needed
- **COM Verification**: Ensures proper API access
- **Calendar Selection**: Targets main calendar (avoids birthday/holiday noise)

## Troubleshooting

### Common Issues

**"Outlook not found"**
```powershell
# Check if Outlook is running
Get-Process outlook -ErrorAction SilentlyContinue
# Start manually if needed
```

**"Authentication failed"**
- Verify iCloud email in `.env`
- Regenerate app-specific password
- Test SMTP connection: `telnet smtp.mail.me.com 465`

**"No events exported"**
- Check date range settings
- Verify Outlook calendar has events
- Run individual export: `python export_outlook_calendar.py`

**"Permission denied"**
- Run PowerShell as Administrator
- Check Windows Defender exclusions
- Verify COM permissions

### Debug Mode
Run individual components to isolate issues:
```powershell
# Test Outlook connection
python outlook_manager.py

# Test export only  
python export_outlook_calendar.py

# Test email only
python email_icloud.py
```

## Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)  
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Why "FreeMyCalendar"?

Born from the frustration of corporate calendar silos, this tool liberates your work schedule to live alongside your personal calendar. Perfect for:

- **Remote Workers**: Keep work/personal calendars separate but visible
- **Consultants**: Sync multiple client calendars  
- **Students**: Academic calendar integration
- **Anyone**: Who wants their calendar on every device

## Acknowledgments

- Microsoft Outlook COM interface integration
- iCloud SMTP compatibility  
- Python `win32com` for seamless Windows integration
- Built for the modern hybrid work environment

---

*Made with ❤️ for calendar freedom*
