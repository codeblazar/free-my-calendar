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
outlook_email = os.getenv("OUTLOOK_EMAIL", "")  # Specific mailbox to access

# Ensure export directory exists
os.makedirs(export_dir, exist_ok=True)
export_path = os.path.join(export_dir, csv_filename)

# Configure date range: 2 weeks in the past, 12 weeks into the future
outlook_start = datetime.now() - timedelta(weeks=2)
outlook_end = datetime.now() + timedelta(weeks=12)

# Connect to Outlook
print("Connecting to Outlook...")
outlook = win32com.client.Dispatch("Outlook.Application")
namespace = outlook.GetNamespace("MAPI")

# Access specific mailbox if configured, otherwise use default
if outlook_email:
    print(f"Accessing mailbox: {outlook_email}")
    try:
        # Try to get the specific mailbox
        recipient = namespace.CreateRecipient(outlook_email)
        recipient.Resolve()
        if recipient.Resolved:
            # Get the main calendar folder (not birthday calendar)
            mailbox = namespace.GetSharedDefaultFolder(recipient, 9)  # 9 = Calendar folder
            print(f"Successfully accessed {outlook_email} main calendar")
            
            # List available calendar folders and find the work calendar
            print("Available calendar folders:")
            work_calendar = None
            try:
                # Get the mailbox store
                store = mailbox.Parent
                print(f"Store name: {store.Name}")
                
                # Look for all calendar folders in this mailbox
                folders = store.Folders
                for folder in folders:
                    if folder.Name == "Calendar":  # This is the Calendar top-level folder
                        print(f"Found Calendar folder with subfolders:")
                        if hasattr(folder, 'Folders'):
                            for subfolder in folder.Folders:
                                if subfolder.DefaultItemType == 1:  # Calendar items
                                    print(f"  - Calendar: {subfolder.Name} ({subfolder.Items.Count} items)")
                        
                        # Use the main Calendar folder itself (not subfolders)
                        # This should be Peter Kenny's primary calendar
                        work_calendar = folder.Items
                        print(f"  >> Using main calendar folder for Peter Kenny")
                        
                        # Verify we're getting reasonable results by checking item count
                        item_count = work_calendar.Count
                        print(f"  >> Calendar contains {item_count} total items")
                        
                        # If the main calendar seems empty or problematic, 
                        # we could add fallback logic here in the future
                        break
                
                # Use work calendar if found, otherwise fall back to default
                if work_calendar:
                    calendar = work_calendar
                else:
                    print("Main calendar not found, using default calendar")
                    calendar = mailbox.Items
                            
            except Exception as e:
                print(f"Could not list folders: {e}")
                calendar = mailbox.Items
                
        else:
            print(f"Could not resolve {outlook_email}, falling back to default calendar")
            calendar = namespace.GetDefaultFolder(9).Items
    except Exception as e:
        print(f"Error accessing {outlook_email}: {e}")
        print("Falling back to default calendar")
        calendar = namespace.GetDefaultFolder(9).Items
else:
    print("Using default calendar")
    calendar = namespace.GetDefaultFolder(9).Items

calendar.IncludeRecurrences = True
calendar.Sort("[Start]", True)  # True for descending (newest first)

# Check total items in calendar first
total_items = calendar.Count
print(f"Total items in calendar: {total_items}")

# Apply date restriction to get current events only
print(f"Applying date filter: {outlook_start.strftime('%Y-%m-%d')} to {outlook_end.strftime('%Y-%m-%d')}")

# Create a restriction to filter events by date range - try different format
start_date_str = outlook_start.strftime('%m/%d/%Y')
end_date_str = outlook_end.strftime('%m/%d/%Y')
restriction = f"[Start] >= '{start_date_str}' AND [Start] < '{end_date_str}'"
print(f"Restriction filter: {restriction}")

try:
    # Apply the restriction
    filtered_items = calendar.Restrict(restriction)
    filtered_count = filtered_items.Count
    print(f"Filtered items count: {filtered_count}")
    
    # If restriction doesn't work properly, let's manually filter first few items
    if filtered_count == 0 or filtered_count > 10000:  # Unreasonable number
        print("Restriction filter may not be working, trying manual approach...")
        print("Checking first 10 items in calendar:")
        item_count = 0
        for item in calendar:
            if item_count >= 10:
                break
            try:
                print(f"  {item.Subject} - {item.Start} - Type: {type(item.Start)}")
                item_count += 1
            except Exception as e:
                print(f"  Error reading item: {e}")
                item_count += 1
    
    # Export to CSV with restricted items
    with open(export_path, "w", newline='', encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Subject", "Start", "End", "Location", "Body"])
        exported_count = 0
        
        # Use filtered items if reasonable, otherwise manual filter
        items_to_process = calendar  # Always use manual filter since restriction isn't working
        
        for item in items_to_process:
            try:
                # Manual date check - only get current events
                item_start = item.Start
                
                # Convert to Python datetime for comparison
                if hasattr(item_start, 'date'):
                    item_date = item_start.date()
                elif isinstance(item_start, str):
                    item_date = datetime.strptime(item_start.split(' ')[0], '%Y-%m-%d').date()
                else:
                    item_date = item_start.date()
                
                # Only process events in our target date range (current month)
                if not (outlook_start.date() <= item_date <= outlook_end.date()):
                    continue
                
                print(f"Processing: {getattr(item, 'Subject', 'No Subject')} - {item.Start}")
                
                # Get event details with better error handling
                subject = getattr(item, 'Subject', 'No Subject')
                start_time = item.Start
                end_time = getattr(item, 'End', '')
                location = getattr(item, 'Location', '')
                body = getattr(item, 'Body', '')
                
                # Clean up body text
                if body:
                    body = str(body).replace('\n', ' ').replace('\r', ' ')[:body_char_limit]
                
                writer.writerow([
                    subject,
                    start_time,
                    end_time,
                    location,
                    body
                ])
                exported_count += 1
                
                # No limit - get all events in the date range
                if exported_count >= 100:  # Safety limit
                    print("Reached 100 events limit...")
                    break
                    
            except Exception as e:
                print(f"Skipping an item due to error: {e}")
                continue
                
except Exception as e:
    print(f"Error applying restriction: {e}")
    # Fallback to manual filtering
    with open(export_path, "w", newline='', encoding="utf-8") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Subject", "Start", "End", "Location", "Body"])
        exported_count = 0
        print("Using manual date filtering...")
        
        for item in calendar:
            try:
                # Check if event starts within our date range
                item_start = item.Start
                if hasattr(item_start, 'date'):
                    item_date = item_start.date()
                else:
                    # Handle different datetime formats
                    if isinstance(item_start, str):
                        item_date = datetime.strptime(item_start.split(' ')[0], '%Y-%m-%d').date()
                    else:
                        item_date = item_start.date()
                
                # Check if item is in our date range
                if outlook_start.date() <= item_date <= outlook_end.date():
                    print(f"Processing: {item.Subject} - {item.Start}")
                    writer.writerow([
                        item.Subject,
                        item.Start,
                        item.End,
                        item.Location if hasattr(item, 'Location') else "",
                        str(item.Body).replace('\n', ' ').replace('\r', ' ')[:body_char_limit] if hasattr(item, 'Body') else ""
                    ])
                    exported_count += 1
                    
                    # Limit for testing
                    if exported_count >= 50:
                        print("Reached 50 events limit for testing...")
                        break
                        
            except Exception as e:
                print(f"Skipping an item due to error: {e}")

print(f"Export complete! Exported {exported_count} events to: {export_path}")