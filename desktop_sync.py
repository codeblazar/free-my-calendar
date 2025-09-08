import os
import sys
from datetime import datetime
from pathlib import Path

# Add the script directory to Python path
script_dir = Path(__file__).parent.absolute()
sys.path.insert(0, str(script_dir))

from sync import run_sync_with_deletions

def main():
    """Ad-hoc calendar sync with user interaction"""
    print("🗓️  Outlook Calendar Sync - Manual Run")
    print("=" * 50)
    print(f"📅 Date: {datetime.now().strftime('%A, %B %d, %Y')}")
    print(f"⏰ Time: {datetime.now().strftime('%I:%M %p')}")
    print()
    
    print("This will:")
    print("  1. ✅ Check/Start Classic Outlook")
    print("  2. 📤 Export calendar events (2 weeks past → 12 weeks future)")
    print("  3. 🔄 Track deletions and modifications")
    print("  4. 📧 Email calendar to your Mac (maxgroove@me.com)")
    print()
    
    # Ask user if they want to proceed
    while True:
        choice = input("Do you want to continue? (y/n): ").lower().strip()
        if choice in ['y', 'yes']:
            break
        elif choice in ['n', 'no']:
            print("Sync cancelled by user.")
            input("Press Enter to exit...")
            return
        else:
            print("Please enter 'y' for yes or 'n' for no.")
    
    print()
    print("🚀 Starting sync process...")
    print("-" * 50)
    
    try:
        success = run_sync_with_deletions()
        
        print("-" * 50)
        if success:
            print("✅ Calendar sync completed successfully!")
            print()
            print("📧 Check your email inbox (maxgroove@me.com) for:")
            print("   • Main calendar file (outlook_calendar_export.ics)")
            print("   • Deletion commands (if any events were removed)")
            print()
            print("📱 Import the files into your Mac calendar in this order:")
            print("   1. Import deletion file first (if received)")
            print("   2. Import main calendar file second")
        else:
            print("❌ Calendar sync failed!")
            print("Check the error messages above for details.")
            print()
            print("💡 Common issues:")
            print("   • Outlook not running or not signed in")
            print("   • Network connectivity problems")
            print("   • Invalid email credentials")
    
    except KeyboardInterrupt:
        print("\n⏹️  Sync interrupted by user (Ctrl+C)")
    
    except Exception as e:
        print(f"\n❌ Unexpected error: {e}")
        print("Please check your configuration and try again.")
    
    print()
    input("Press Enter to exit...")

if __name__ == "__main__":
    main()
