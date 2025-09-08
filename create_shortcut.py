import os
import win32com.client
from pathlib import Path

def create_desktop_shortcut():
    """Create a desktop shortcut for calendar sync"""
    
    # Get paths - try OneDrive Desktop first, then fallback to regular Desktop
    script_dir = Path(__file__).parent.absolute()
    
    # Try OneDrive Desktop path first
    onedrive_desktop = Path(r"C:\Users\peter_kenny\OneDrive - Republic Polytechnic\Desktop")
    regular_desktop = Path.home() / "Desktop"
    
    if onedrive_desktop.exists():
        desktop_path = onedrive_desktop
        print(f"Using OneDrive Desktop: {desktop_path}")
    elif regular_desktop.exists():
        desktop_path = regular_desktop
        print(f"Using regular Desktop: {desktop_path}")
    else:
        print("❌ Desktop folder not found!")
        print(f"Tried: {onedrive_desktop}")
        print(f"Tried: {regular_desktop}")
        return False
    
    # Shortcut details
    shortcut_name = "Sync Calendar.lnk"
    shortcut_path = desktop_path / shortcut_name
    target_script = script_dir / "desktop_sync.bat"
    
    # Icon path (use calendar icon if available, otherwise batch file icon)
    icon_path = str(target_script)  # Use batch file icon
    
    try:
        # Create shortcut using Windows COM
        shell = win32com.client.Dispatch("WScript.Shell")
        shortcut = shell.CreateShortCut(str(shortcut_path))
        
        # Set shortcut properties
        shortcut.Targetpath = str(target_script)
        shortcut.Arguments = ""  # No arguments needed for batch file
        shortcut.WorkingDirectory = str(script_dir)
        shortcut.Description = "Outlook Calendar Sync - Manual execution with user interface"
        shortcut.IconLocation = f"{icon_path},0"
        
        # Save the shortcut
        shortcut.save()
        
        print(f"✅ Desktop shortcut created successfully!")
        print(f"📍 Location: {shortcut_path}")
        print(f"🎯 Target: {target_script}")
        print()
        print("You can now double-click the '📅 Sync Calendar' icon on your desktop")
        print("to run calendar sync manually anytime!")
        
        return True
        
    except Exception as e:
        print(f"❌ Failed to create desktop shortcut: {e}")
        print()
        print("Manual shortcut creation:")
        print(f"1. Right-click on desktop → New → Shortcut")
        print(f"2. Target: {target_script}")
        print(f"3. Start in: {script_dir}")
        print(f"4. Name: Sync Calendar")
        
        return False

if __name__ == "__main__":
    print("🔗 Creating Desktop Shortcut for Calendar Sync")
    print("=" * 50)
    create_desktop_shortcut()
