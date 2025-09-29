import imaplib
import email
from email import policy
import os
import re
import sys
import datetime

# Configuration from environment variables
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.gmail.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
TARGET_SENDER = os.environ.get("TARGET_SENDER", "paulo.costa@dlh.de")

if not all([EMAIL_USER, EMAIL_PASS, IMAP_SERVER, TARGET_SENDER]):
    sys.exit("ERROR: EMAIL_USER, EMAIL_PASS, IMAP_SERVER, and TARGET_SENDER must be set.")

# Directory to save attachments
ATTACH_DIR = os.path.join(os.getcwd(), "attachments")
os.makedirs(ATTACH_DIR, exist_ok=True)

def sanitize_filename(filename):
    """Remove invalid characters from filenames"""
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

print("Connecting to Gmail IMAP server...")
try:
    with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as mail:
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        # Search emails from target sender
        status, messages = mail.search(None, f'(FROM "{TARGET_SENDER}")')
        if status != "OK":
            sys.exit("Failed to search emails.")

        msg_ids = messages[0].split()
        if not msg_ids:
            print(f"No messages found from {TARGET_SENDER}.")
            sys.exit(0)

        # Process the newest message
        newest_msg_id = msg_ids[-1]
        print(f"Processing newest message ID: {newest_msg_id.decode()} from {TARGET_SENDER}")

        status, data = mail.fetch(newest_msg_id, "(RFC822)")
        if status != "OK" or not data:
            sys.exit("Failed to fetch the newest message.")

        raw_email = None
        for item in data:
            if isinstance(item, tuple) and item[1]:
                raw_email = item[1]
                break

        if not raw_email:
            sys.exit("No email bytes found for the newest message.")

        msg = email.message_from_bytes(raw_email, policy=policy.default)
        subject = msg.get("subject", "(no subject)")
        print(f"Subject: {subject}")

        found_any = False
        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                filename = sanitize_filename(filename)
                timestamp = datetime.datetime.now().strftime("%Y%m%d%H%M%S")
                filepath = os.path.join(ATTACH_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {filepath}")
                found_any = True

        if not found_any:
            print("No attachments found in the newest message.")

        mail.logout()

except imaplib.IMAP4.error as e:
    sys.exit(f"IMAP error: {e}")
