import schedule
import time
import subprocess
import os
import logging
from datetime import datetime

# Setup logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('calendar_sync_scheduler.log'),
        logging.StreamHandler()
    ]
)

def run_calendar_sync():
    """Run the complete calendar sync process"""
    try:
        logging.info("Starting calendar sync process")
        
        # Change to the correct directory
        os.chdir(r"C:\OutlookCalendarExports")
        
        # Step 1: Export from Outlook
        logging.info("Step 1: Exporting from Outlook")
        result = subprocess.run(['python', 'export_outlook_calendar.py'], 
                              capture_output=True, text=True, timeout=300)
        if result.returncode != 0:
            raise Exception(f"Export failed: {result.stderr}")
        
        # Step 2: Convert to ICS
        logging.info("Step 2: Converting to ICS")
        result = subprocess.run(['python', 'csv_to_ics.py'], 
                              capture_output=True, text=True, timeout=60)
        if result.returncode != 0:
            raise Exception(f"Conversion failed: {result.stderr}")
        
        # Step 3: Email the file
        logging.info("Step 3: Emailing ICS file")
        result = subprocess.run(['python', 'email_icloud.py'], 
                              capture_output=True, text=True, timeout=120)
        if result.returncode != 0:
            raise Exception(f"Email failed: {result.stderr}")
        
        logging.info("Calendar sync completed successfully")
        
    except Exception as e:
        logging.error(f"Calendar sync failed: {str(e)}")

# Schedule the sync twice daily
schedule.every().day.at("12:00").do(run_calendar_sync)
schedule.every().day.at("21:00").do(run_calendar_sync)  # 9 PM

logging.info("Calendar sync scheduler started. Running at 12:00 PM and 9:00 PM daily.")
logging.info("Press Ctrl+C to stop...")

try:
    while True:
        schedule.run_pending()
        time.sleep(60)  # Check every minute
except KeyboardInterrupt:
    logging.info("Scheduler stopped by user")
