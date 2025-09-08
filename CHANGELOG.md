# Changelog

All notable changes to FreeMyCalendar will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.0.0] - 2025-09-08

### 🚀 Major Features Added
- **Deletion Tracking System**: Full sync with tracking of deleted events
- **Smart Outlook Management**: Auto-detection and startup of Classic Outlook
- **Desktop Integration**: One-click sync via desktop shortcut
- **Enhanced Date Range**: Configurable past/future sync windows (default: 2 weeks past, 12 weeks future)
- **Improved User Experience**: Interactive sync with progress displays

### 🔧 Technical Improvements
- **sync.py**: New main orchestrator replacing individual script execution
- **outlook_manager.py**: Process detection and management system
- **sync_tracker.py**: MD5-based event tracking for proper deletion handling
- **desktop_sync.py**: User-friendly interactive interface
- **create_shortcut.py**: Automated desktop shortcut creation

### 🐛 Bug Fixes
- Fixed calendar folder detection (was accessing birthday calendar instead of main calendar)
- Resolved virtual environment issues in batch scripts
- Improved error handling for COM interface failures
- Fixed date range filtering for proper sync boundaries

### 🗂️ Project Structure
- Reorganized codebase into logical components
- Cleaned up temporary and development files
- Enhanced documentation with complete setup guide
- Added comprehensive troubleshooting section

### 📱 User Interface
- Desktop shortcut for one-click synchronization
- Interactive prompts with progress feedback
- Startup folder integration option
- Multiple execution methods (GUI, command line, automated)

### ⚙️ Configuration
- Simplified .env setup with sensible defaults
- Enhanced environment variable documentation
- Flexible SMTP and calendar settings
- Corporate environment compatibility

### 🔒 Security
- Removed personal credentials from repository
- Enhanced .gitignore for sensitive files
- App-specific password integration guide
- Secure SMTP authentication

---

## [1.0.0] - 2025-08-30

### Added
- Initial release of Outlook Calendar Export Tool
- Export Outlook calendar events to CSV format
- Convert CSV to iCalendar (ICS) format with unique event IDs
- Email ICS files to iCloud account via SMTP
- Environment variable configuration system
- Windows Task Scheduler automation support
- Batch script for automated execution
- PowerShell alternative script
- Python scheduler option
- Comprehensive error handling and logging
- Support for recurring events
- Configurable export duration and file paths
- Smart sync with unique event IDs to prevent duplicates
- Proper calendar metadata for better app compatibility

### Features
- Three-step export process: Outlook → CSV → ICS → Email
- Automated twice-daily scheduling (12PM & 9PM)
- Configurable via `.env` file
- Support for separate work/personal calendars
- Handles event updates and deletions properly
- Cross-platform Python scripts (Windows-focused)
- Detailed documentation and troubleshooting guide

### Dependencies
- Python 3.7+
- pywin32 (Windows COM interface)
- python-dotenv (environment variables)
