import os
import socket
from dotenv import load_dotenv
import smtplib
from email.message import EmailMessage

# Load environment variables
load_dotenv()
print("Loaded environment variables.")

# Configuration from environment variables
sender_email = os.getenv("ICLOUD_EMAIL")  # Use same email as authenticated account
receiver_email = os.getenv("ICLOUD_EMAIL")
subject = os.getenv("EMAIL_SUBJECT", "Automated Outlook Calendar Export")
body = os.getenv("EMAIL_BODY", "Find attached the latest Outlook calendar export as iCal.")

# File paths
export_dir = os.getenv("EXPORT_DIRECTORY", r"C:\OutlookCalendarExports")
ics_filename = os.getenv("ICS_FILENAME", "outlook_calendar_export.ics")
ics_path = os.path.join(export_dir, ics_filename)

# SMTP configuration from environment
smtp_server = os.getenv("SMTP_SERVER", "smtp.mail.me.com")
smtp_port = int(os.getenv("SMTP_PORT", 465))
smtp_timeout = int(os.getenv("SMTP_TIMEOUT", 30))
smtp_user = os.getenv("ICLOUD_EMAIL")
smtp_password = os.getenv("ICLOUD_APP_PASSWORD")

print(f"Preparing to send email from {sender_email} to {receiver_email}")

# Validate environment variables
if not smtp_user or not smtp_password:
    print("Error: ICLOUD_EMAIL or ICLOUD_APP_PASSWORD not set in .env file")
    exit(1)

# Check if ICS file exists
if not os.path.exists(ics_path):
    print(f"Error: ICS file not found at {ics_path}")
    print("Please run csv_to_ics.py first to create the ICS file.")
    exit(1)

print("Opening ICS file...")

# Create email message
msg = EmailMessage()
msg["From"] = sender_email
msg["To"] = receiver_email
msg["Subject"] = subject
msg.set_content(body)

# Attach ICS file
with open(ics_path, "rb") as f:
    print("ICS file attached.")
    msg.add_attachment(f.read(), maintype="text", subtype="calendar", filename="outlook_calendar_export.ics")

# Send email
print(f"Connecting to SMTP server: {smtp_server}:{smtp_port}")
try:
    with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=smtp_timeout) as server:
        print("Connected to SMTP server.")
        print(f"Logging in as {smtp_user}...")
        server.login(smtp_user, smtp_password)
        print("Logged in successfully. Sending email...")
        server.send_message(msg)
        print("Email sent successfully!")
        
except socket.timeout:
    print("Error: Connection timeout. Check your internet connection.")
except ConnectionRefusedError:
    print("Error: Connection refused. Check SMTP server and port.")
except smtplib.SMTPAuthenticationError:
    print("Error: Authentication failed. Check your iCloud email and app password.")
except smtplib.SMTPException as e:
    print(f"SMTP error: {e}")
except Exception as e:
    print(f"Unexpected error: {e}")

print("Process completed.")