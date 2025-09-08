import os
import subprocess
import sys
from datetime import datetime
from dotenv import load_dotenv
from sync_tracker import SyncTracker
from outlook_manager import OutlookManager

# Load environment variables
load_dotenv()

def run_sync_with_deletions():
    """Run the complete sync process with deletion tracking"""
    print(f"Starting Outlook Calendar Sync at {datetime.now()}")
    print("=" * 60)
    
    # Step 0: Ensure Classic Outlook is running
    print("Step 0: Verifying Classic Outlook...")
    outlook_manager = OutlookManager()
    if not outlook_manager.ensure_classic_outlook_running():
        print("❌ Failed to start or verify Classic Outlook")
        print("Please ensure Classic Outlook is installed and accessible.")
        return False
    
    print("✅ Classic Outlook is ready")
    print()
    
    # Initialize sync tracker
    tracker = SyncTracker()
    
    # Step 1: Load previous sync data
    print("Step 1: Loading previous sync data...")
    previous_data = tracker.load_previous_sync()
    if previous_data:
        print(f"  Previous sync: {previous_data.get('sync_date', 'Unknown')}")
        print(f"  Previous events: {previous_data.get('total_events', 0)}")
    else:
        print("  No previous sync data found (first run)")
    
    # Step 2: Export from Outlook
    print("\nStep 2: Exporting from Outlook...")
    try:
        result = subprocess.run([sys.executable, "export_outlook_calendar.py"], 
                              capture_output=True, text=True, check=True)
        print("  Export completed successfully")
    except subprocess.CalledProcessError as e:
        print(f"  Export failed: {e}")
        return False
    
    # Step 3: Load current events and compare
    print("\nStep 3: Analyzing changes...")
    csv_file = os.path.join(os.getenv("EXPORT_DIRECTORY", "."), 
                           os.getenv("CSV_FILENAME", "outlook_calendar_export.csv"))
    
    tracker.load_current_events(csv_file)
    added, deleted, modified = tracker.find_changes()
    
    print(f"  Current events: {len(tracker.current_events)}")
    print(f"  Added events: {len(added)}")
    print(f"  Deleted events: {len(deleted)}")
    print(f"  Modified events: {len(modified)}")
    
    # Handle modified events - add old IDs to deletion list
    deletion_ids = deleted.copy()
    if modified:
        print("  Modified events details:")
        for old_id, new_id in modified:
            if old_id in tracker.previous_events and new_id in tracker.current_events:
                old_event = tracker.previous_events[old_id]
                new_event = tracker.current_events[new_id]
                print(f"    - Modified: {old_event['subject']}")
                print(f"      Old: {old_event['start']} to {old_event['end']}")
                print(f"      New: {new_event['start']} to {new_event['end']}")
                deletion_ids.append(old_id)  # Add old version to deletion list
    
    # Show details of deletions
    if deletion_ids:
        print("  Events to be deleted:")
        for event_id in deletion_ids[:5]:  # Show first 5
            if event_id in tracker.previous_events:
                event = tracker.previous_events[event_id]
                print(f"    - {event['subject']} ({event['start']} to {event['end']})")
        if len(deletion_ids) > 5:
            print(f"    ... and {len(deletion_ids) - 5} more")
    
    # Step 4: Convert to ICS
    print("\nStep 4: Converting to ICS format...")
    try:
        result = subprocess.run([sys.executable, "csv_to_ics.py"], 
                              capture_output=True, text=True, check=True)
        print("  ICS conversion completed")
    except subprocess.CalledProcessError as e:
        print(f"  ICS conversion failed: {e}")
        return False
    
    # Step 5: Create deletion ICS if needed
    ics_file = os.path.join(os.getenv("EXPORT_DIRECTORY", "."), 
                           os.getenv("ICS_FILENAME", "outlook_calendar_export.ics"))
    deletion_file = None
    
    if deletion_ids:
        print(f"\nStep 5: Creating deletion ICS file for {len(deletion_ids)} deleted/old events...")
        deletion_file = tracker.generate_deletion_ics(deletion_ids, ics_file)
        if deletion_file:
            print(f"  Deletion file created: {deletion_file}")
    else:
        print("\nStep 5: No deletions to process")
    
    # Step 6: Email the calendar file
    print("\nStep 6: Sending calendar via email...")
    try:
        result = subprocess.run([sys.executable, "email_icloud.py"], 
                              capture_output=True, text=True, check=True)
        print("  Main calendar emailed successfully")
    except subprocess.CalledProcessError as e:
        print(f"  Email failed: {e}")
        return False
    
    # Step 7: Email deletion file if it exists
    if deletion_file and os.path.exists(deletion_file):
        print("\nStep 7: Sending deletion commands via email...")
        try:
            # Temporarily modify the environment to send the deletion file
            original_ics = os.getenv("ICS_FILENAME")
            os.environ["ICS_FILENAME"] = os.path.basename(deletion_file)
            os.environ["EMAIL_SUBJECT"] = "Calendar Event Deletions - Pete Work"
            os.environ["EMAIL_BODY"] = f"Deletion commands for {len(deletion_ids)} removed/modified calendar events. Import this FIRST to remove old versions, then import the main calendar."
            
            result = subprocess.run([sys.executable, "email_icloud.py"], 
                                  capture_output=True, text=True, check=True)
            print("  Deletion commands emailed successfully")
            
            # Restore original environment
            if original_ics:
                os.environ["ICS_FILENAME"] = original_ics
                
        except subprocess.CalledProcessError as e:
            print(f"  Deletion email failed: {e}")
    else:
        print("\nStep 7: No deletion file to send")
    
    # Step 8: Save current sync data for next time
    print("\nStep 8: Saving sync tracking data...")
    tracker.save_current_sync()
    
    print("\n" + "=" * 60)
    print("Enhanced Calendar Sync completed successfully!")
    print(f"Summary:")
    print(f"  - Total events synced: {len(tracker.current_events)}")
    print(f"  - New events: {len(added)}")
    print(f"  - Modified events: {len(modified)}")
    print(f"  - Deleted events: {len(deleted)}")
    print(f"  - Total deletions sent: {len(deletion_ids) if 'deletion_ids' in locals() else 0}")
    print(f"  - Files sent: {'2 (calendar + deletions)' if deletion_file else '1 (calendar only)'}")
    
    return True

if __name__ == "__main__":
    success = run_sync_with_deletions()
    if not success:
        sys.exit(1)
