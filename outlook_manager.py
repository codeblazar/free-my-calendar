import subprocess
import time
import win32com.client
import psutil
import os
from pathlib import Path

class OutlookManager:
    def __init__(self):
        self.classic_outlook_paths = [
            r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE",
            r"C:\Program Files (x86)\Microsoft Office\root\Office16\OUTLOOK.EXE",
            r"C:\Program Files\Microsoft Office\Office16\OUTLOOK.EXE",
            r"C:\Program Files (x86)\Microsoft Office\Office16\OUTLOOK.EXE",
            # Add more common paths as needed
        ]
    
    def is_outlook_running(self):
        """Check if any version of Outlook is running"""
        for proc in psutil.process_iter(['pid', 'name']):
            try:
                if proc.info['name'].lower() == 'outlook.exe':
                    return True, proc.info['pid']
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False, None
    
    def is_classic_outlook_running(self):
        """Check if classic Outlook is running (not new Outlook)"""
        for proc in psutil.process_iter(['pid', 'name', 'exe']):
            try:
                if proc.info['name'].lower() == 'outlook.exe':
                    exe_path = proc.info.get('exe', '')
                    # New Outlook typically runs from WindowsApps folder
                    if 'WindowsApps' not in exe_path and 'Microsoft.OutlookForWindows' not in exe_path:
                        return True, proc.info['pid']
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                continue
        return False, None
    
    def find_classic_outlook_path(self):
        """Find the path to classic Outlook executable"""
        for path in self.classic_outlook_paths:
            if os.path.exists(path):
                return path
        
        # Try to find it in registry or common locations
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                               r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE")
            path, _ = winreg.QueryValueEx(key, "")
            winreg.CloseKey(key)
            if os.path.exists(path):
                return path
        except:
            pass
        
        return None
    
    def start_classic_outlook(self):
        """Start classic Outlook if it's not running"""
        outlook_path = self.find_classic_outlook_path()
        if not outlook_path:
            raise Exception("Classic Outlook executable not found. Please install Microsoft Office.")
        
        print(f"Starting Classic Outlook from: {outlook_path}")
        try:
            subprocess.Popen([outlook_path])
            return True
        except Exception as e:
            raise Exception(f"Failed to start Classic Outlook: {e}")
    
    def wait_for_outlook_com(self, timeout=30):
        """Wait for Outlook COM interface to be available"""
        start_time = time.time()
        while time.time() - start_time < timeout:
            try:
                outlook = win32com.client.Dispatch("Outlook.Application")
                namespace = outlook.GetNamespace("MAPI")
                # Try to access a folder to ensure it's fully loaded
                namespace.GetDefaultFolder(6)  # Inbox folder
                print("Outlook COM interface is ready")
                return True
            except Exception as e:
                print(f"Waiting for Outlook COM interface... ({e})")
                time.sleep(2)
        
        return False
    
    def ensure_classic_outlook_running(self):
        """Ensure Classic Outlook is running and COM interface is ready"""
        print("Checking Outlook status...")
        
        # Check if classic Outlook is already running
        is_classic_running, pid = self.is_classic_outlook_running()
        
        if is_classic_running:
            print(f"Classic Outlook is already running (PID: {pid})")
        else:
            # Check if any Outlook is running (might be new Outlook)
            is_any_running, pid = self.is_outlook_running()
            if is_any_running:
                print(f"WARNING: New Outlook or other Outlook version detected (PID: {pid})")
                print("This tool requires Classic Outlook. Please close New Outlook and restart this script.")
                choice = input("Do you want to continue anyway? (y/n): ").lower()
                if choice != 'y':
                    return False
            
            # Start classic Outlook
            print("Starting Classic Outlook...")
            try:
                self.start_classic_outlook()
                print("Waiting for Classic Outlook to initialize...")
                time.sleep(5)  # Give Outlook time to start
            except Exception as e:
                print(f"Error starting Classic Outlook: {e}")
                return False
        
        # Wait for COM interface to be ready
        print("Verifying Outlook COM interface...")
        if self.wait_for_outlook_com():
            print("✅ Classic Outlook is ready for calendar export")
            return True
        else:
            print("❌ Outlook COM interface not available")
            return False

def main():
    """Test the Outlook manager"""
    manager = OutlookManager()
    success = manager.ensure_classic_outlook_running()
    if success:
        print("Outlook is ready for calendar operations!")
    else:
        print("Failed to prepare Outlook for calendar operations.")
        return False
    return True

if __name__ == "__main__":
    main()
