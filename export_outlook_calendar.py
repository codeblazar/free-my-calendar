import csv
import win32com.client
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get configuration from environment
export_dir = os.getenv("EXPORT_DIRECTORY", r"C:\OutlookCalendarExports")
csv_filename = os.getenv("CSV_FILENAME", "outlook_calendar_export.csv")
export_days = int(os.getenv("EXPORT_DAYS", 30))
body_char_limit = int(os.getenv("BODY_CHAR_LIMIT", 500))

# Ensure export directory exists
os.makedirs(export_dir, exist_ok=True)
export_path = os.path.join(export_dir, csv_filename)

# Configure date range
outlook_start = datetime.now()
outlook_end = outlook_start + timedelta(days=export_days)

# Connect to Outlook
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")
calendar = namespace.GetDefaultFolder(9).Items
calendar.IncludeRecurrences = True
calendar.Sort("[Start]")

# Restrict to date range
restriction = "[Start] >= '{}' AND [End] <= '{}'".format(
    outlook_start.strftime('%m/%d/%Y %H:%M %p'),
    outlook_end.strftime('%m/%d/%Y %H:%M %p')
)
restricted_items = calendar.Restrict(restriction)

# Export to CSV
with open(export_path, "w", newline='', encoding="utf-8") as csvfile:
    writer = csv.writer(csvfile)
    writer.writerow(["Subject", "Start", "End", "Location", "Body"])
    for item in restricted_items:
        try:
            writer.writerow([
                item.Subject,
                item.Start,
                item.End,
                item.Location,
                str(item.Body).replace('\n', ' ').replace('\r', ' ')[:body_char_limit]  # Limit body size
            ])
        except Exception as e:
            print(f"Skipping an item due to error: {e}")

print(f"Export complete! File: {export_path}")