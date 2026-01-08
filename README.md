# CalendarFree

Sync Outlook (Office 365 desktop) calendar entries into an iCloud calendar via CalDAV by fully overwriting the target calendar each run.

## Setup
- Copy `.env.example` to `.env` and set values:
  - `DELIVERY_METHOD` must be `caldav`.
  - `OUTLOOK_STORE_NAME` (e.g., your mailbox address) and `OUTLOOK_CALENDAR_NAME` (e.g., `Calendar`); both help target the correct Outlook calendar.
  - `SYNC_DAYS_PAST` / `SYNC_DAYS_FUTURE` to define the window exported from Outlook (defaults: past 3 days, next 60 days).
  - `TIMEZONE` defaults to `Asia/Singapore`.
  - `ICLOUD_EMAIL`, `ICLOUD_APP_PASSWORD` (app-specific password from appleid.apple.com), and `ICLOUD_CALDAV_CALENDAR_NAME` (must match the iCloud calendar display name).
- Create and activate a virtual environment, then install dependencies:
  - `python -m venv .venv`
  - Windows PowerShell: `.\.venv\Scripts\Activate.ps1`
  - Install: `python -m pip install -r requirements.txt`

## Usage
- Run the sync: `python sync.py`.
- The script reads events from Outlook within the configured window, deletes all existing events in the specified iCloud calendar, then uploads the Outlook events to iCloud.

## Optional: schedule every 15 minutes (no admin)
- Open PowerShell (normal user) in the project folder and run:
  ```
  $now = Get-Date
  $offset = (15 - ($now.Minute % 15)) % 15
  if ($offset -eq 0) { $offset = 15 }
  $start = $now.AddMinutes($offset).AddSeconds(-$now.Second).AddMilliseconds(-$now.Millisecond)
  $cmd = "C:\Projects\CalendarFree\.venv\Scripts\pythonw.exe"
  $args = "C:\Projects\CalendarFree\sync.py"
  $trigger = New-ScheduledTaskTrigger -Once -At $start -RepetitionInterval (New-TimeSpan -Minutes 15)
  $action = New-ScheduledTaskAction -Execute $cmd -Argument $args -WorkingDirectory "C:\Projects\CalendarFree"
  $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited
  Register-ScheduledTask -TaskName "CalendarFreeSync" -Action $action -Trigger $trigger -Principal $principal -Description "Sync Outlook Calendar to iCloud via CalendarFree (pythonw, hidden)" -Force
  ```
- This schedules a hidden run every 15 minutes (on the quarter-hour boundaries) while you are logged in.

## Notes
- Recurring meetings are expanded within the sync window; attachments are not uploaded.
- The iCloud calendar should be dedicated to this sync since it is fully replaced each run.
