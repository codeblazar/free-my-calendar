import json
import os
import hashlib
from datetime import datetime
from typing import Dict, List, Set, Tuple

class SyncTracker:
    def __init__(self, tracking_file="sync_history.json"):
        self.tracking_file = tracking_file
        self.current_events = {}
        self.previous_events = {}
        
    def load_previous_sync(self) -> Dict:
        """Load the previous sync data from file"""
        if os.path.exists(self.tracking_file):
            try:
                with open(self.tracking_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.previous_events = data.get('events', {})
                    return data
            except Exception as e:
                print(f"Error loading previous sync data: {e}")
                return {}
        return {}
    
    def generate_event_id(self, subject: str, start_time: str, end_time: str) -> str:
        """Generate a unique ID for an event based on its key properties"""
        # Create a hash of subject + start time + end time for unique identification
        event_string = f"{subject}|{start_time}|{end_time}"
        return hashlib.md5(event_string.encode('utf-8')).hexdigest()
    
    def load_current_events(self, csv_file: str) -> Dict:
        """Load current events from CSV export"""
        import csv
        
        self.current_events = {}
        
        try:
            with open(csv_file, 'r', encoding='utf-8') as csvfile:
                reader = csv.DictReader(csvfile)
                for row in reader:
                    event_id = self.generate_event_id(
                        row['Subject'], 
                        row['Start'], 
                        row['End']
                    )
                    self.current_events[event_id] = {
                        'subject': row['Subject'],
                        'start': row['Start'],
                        'end': row['End'],
                        'location': row['Location'],
                        'body': row['Body']
                    }
        except Exception as e:
            print(f"Error loading current events: {e}")
            
        return self.current_events
    
    def find_changes(self) -> Tuple[List[str], List[str], List[str]]:
        """
        Compare current events with previous events
        Returns: (added_event_ids, deleted_event_ids, modified_event_ids)
        """
        current_ids = set(self.current_events.keys())
        previous_ids = set(self.previous_events.keys())
        
        # Find added events (in current but not in previous)
        added = list(current_ids - previous_ids)
        
        # Find deleted events (in previous but not in current)
        deleted = list(previous_ids - current_ids)
        
        # For modified events, we need to be smarter about matching
        # Look for events with same subject but different start/end times
        modified = []
        subjects_current = {self.current_events[eid]['subject']: eid for eid in current_ids}
        subjects_previous = {self.previous_events[eid]['subject']: eid for eid in previous_ids}
        
        # Find events with same subject that have different IDs (meaning time/date changed)
        for subject in subjects_current:
            if subject in subjects_previous:
                current_id = subjects_current[subject]
                previous_id = subjects_previous[subject]
                
                if current_id != previous_id:  # Same subject, different ID = modified
                    # Remove from added/deleted lists since this is a modification
                    if current_id in added:
                        added.remove(current_id)
                    if previous_id in deleted:
                        deleted.remove(previous_id)
                    
                    # Add to modified list - we'll delete the old and add the new
                    modified.append((previous_id, current_id))
        
        return added, deleted, modified
    
    def save_current_sync(self):
        """Save current sync data for next comparison"""
        sync_data = {
            'sync_date': datetime.now().isoformat(),
            'events': self.current_events,
            'total_events': len(self.current_events)
        }
        
        try:
            with open(self.tracking_file, 'w', encoding='utf-8') as f:
                json.dump(sync_data, f, indent=2, ensure_ascii=False)
            print(f"Sync tracking data saved to {self.tracking_file}")
        except Exception as e:
            print(f"Error saving sync data: {e}")
    
    def generate_deletion_ics(self, deleted_event_ids: List[str], output_file: str):
        """Generate an ICS file with deletion commands for removed events"""
        if not deleted_event_ids:
            return
            
        ics_content = []
        ics_content.append("BEGIN:VCALENDAR")
        ics_content.append("VERSION:2.0")
        ics_content.append("PRODID:-//OutlookSync//Calendar Deletion//EN")
        ics_content.append("CALSCALE:GREGORIAN")
        ics_content.append("METHOD:CANCEL")  # This indicates event cancellation/deletion
        ics_content.append("X-WR-CALNAME:Pete Work - Deletions")
        
        for event_id in deleted_event_ids:
            if event_id in self.previous_events:
                event = self.previous_events[event_id]
                
                # Create cancellation event
                ics_content.append("BEGIN:VEVENT")
                ics_content.append(f"UID:{event_id}@outlooksync.local")
                ics_content.append(f"DTSTAMP:{datetime.now().strftime('%Y%m%dT%H%M%SZ')}")
                ics_content.append("STATUS:CANCELLED")
                ics_content.append(f"SUMMARY:{event['subject']}")
                
                # Convert datetime format for ICS
                try:
                    start_dt = datetime.fromisoformat(event['start'].replace('Z', '+00:00'))
                    ics_content.append(f"DTSTART:{start_dt.strftime('%Y%m%dT%H%M%SZ')}")
                    
                    end_dt = datetime.fromisoformat(event['end'].replace('Z', '+00:00'))
                    ics_content.append(f"DTEND:{end_dt.strftime('%Y%m%dT%H%M%SZ')}")
                except:
                    # Fallback for date format issues
                    pass
                
                ics_content.append("END:VEVENT")
        
        ics_content.append("END:VCALENDAR")
        
        # Save deletion ICS file
        deletion_file = output_file.replace('.ics', '_deletions.ics')
        try:
            with open(deletion_file, 'w', encoding='utf-8') as f:
                f.write('\n'.join(ics_content))
            print(f"Deletion ICS file created: {deletion_file}")
            return deletion_file
        except Exception as e:
            print(f"Error creating deletion file: {e}")
            return None

if __name__ == "__main__":
    # Test the sync tracker
    tracker = SyncTracker()
    tracker.load_previous_sync()
    tracker.load_current_events("outlook_calendar_export.csv")
    
    added, deleted, modified = tracker.find_changes()
    
    print(f"Sync Analysis:")
    print(f"  Added events: {len(added)}")
    print(f"  Deleted events: {len(deleted)}")
    print(f"  Modified events: {len(modified)}")
    
    if deleted:
        print(f"  Creating deletion ICS file...")
        tracker.generate_deletion_ics(deleted, "outlook_calendar_export.ics")
    
    tracker.save_current_sync()
