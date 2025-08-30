# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
- selenium (future web automation)
- schedule (Python scheduling alternative)
