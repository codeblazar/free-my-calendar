import csv
from datetime import datetime
import os
import hashlib
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Get configuration from environment
export_dir = os.getenv("EXPORT_DIRECTORY", r"C:\OutlookCalendarExports")
csv_filename = os.getenv("CSV_FILENAME", "outlook_calendar_export.csv")
ics_filename = os.getenv("ICS_FILENAME", "outlook_calendar_export.ics")
calendar_name = os.getenv("CALENDAR_NAME", "Outlook Work Calendar")
calendar_description = os.getenv("CALENDAR_DESCRIPTION", "Exported from Microsoft Outlook")

def generate_event_uid(subject, start_time, location=""):
    """Generate a unique identifier for each event based on its content"""
    # Create a unique string from event details
    unique_string = f"{subject}_{start_time}_{location}".lower()
    # Generate a hash for the UID
    uid_hash = hashlib.md5(unique_string.encode()).hexdigest()
    return f"{uid_hash}@outlook-export"

def csv_to_ics(csv_file, ics_file):
    with open(csv_file, newline='', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        events = list(reader)

    def format_ics_datetime(dt_str):
        # Try additional ISO formats
        for fmt in (
            "%Y-%m-%d %H:%M:%S%z",    # e.g., 2025-09-02 13:30:00+00:00
            "%Y-%m-%d %H:%M:%S",      # e.g., 2025-09-02 13:30:00
            "%m/%d/%Y %I:%M:%S %p",   # fallback, if others exist
            "%Y-%m-%dT%H:%M:%S%z"
        ):
            try:
                # Remove colon from timezone if present (to match %z pattern)
                if '+' in dt_str:
                    dt_left, tz = dt_str.split('+')
                    tz = tz.replace(':','')
                    dt_str2 = f"{dt_left}+{tz}"
                else:
                    dt_str2 = dt_str
                dt = datetime.strptime(dt_str2, fmt)
                return dt.strftime('%Y%m%dT%H%M%SZ')
            except Exception:
                continue
        raise ValueError(f"Unrecognized date format: {dt_str}")

    with open(ics_file, 'w', encoding='utf-8') as f:
        # ICS header with calendar replacement method
        f.write("BEGIN:VCALENDAR\n")
        f.write("VERSION:2.0\n")
        f.write("PRODID:-//Outlook Calendar Export//CSV2ICS//EN\n")
        f.write("CALSCALE:GREGORIAN\n")
        f.write("METHOD:PUBLISH\n")
        f.write(f"X-WR-CALNAME:{calendar_name}\n")
        f.write(f"X-WR-CALDESC:{calendar_description}\n")
        
        for event in events:
            try:
                # Generate unique UID for each event
                event_uid = generate_event_uid(
                    event.get('Subject', ''), 
                    event.get('Start', ''), 
                    event.get('Location', '')
                )
                
                # Get current timestamp for created/modified dates
                now_timestamp = datetime.now().strftime('%Y%m%dT%H%M%SZ')
                
                f.write("BEGIN:VEVENT\n")
                f.write(f"UID:{event_uid}\n")
                f.write(f"DTSTAMP:{now_timestamp}\n")
                f.write(f"CREATED:{now_timestamp}\n")
                f.write(f"LAST-MODIFIED:{now_timestamp}\n")
                f.write(f"SUMMARY:{event['Subject']}\n")
                f.write(f"DTSTART:{format_ics_datetime(str(event['Start']))}\n")
                f.write(f"DTEND:{format_ics_datetime(str(event['End']))}\n")
                if event.get('Location'): 
                    f.write(f"LOCATION:{event['Location']}\n")
                if event.get('Body'): 
                    # Clean up description text for ICS format
                    description = str(event['Body']).replace('\n', '\\n').replace(',', '\\,').replace(';', '\\;')
                    f.write(f"DESCRIPTION:{description}\n")
                f.write("STATUS:CONFIRMED\n")
                f.write("TRANSP:OPAQUE\n")
                f.write("END:VEVENT\n")
            except Exception as e:
                print(f"Error processing event: {e}")
        f.write("END:VCALENDAR\n")
    print(f"Done! File saved as: {ics_file}")

if __name__ == '__main__':
    csv_path = os.path.join(export_dir, csv_filename)
    ics_path = os.path.join(export_dir, ics_filename)
    csv_to_ics(csv_path, ics_path)
