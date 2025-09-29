import imaplib
import email
import os
from email.utils import parsedate_to_datetime

EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

ATTACH_DIR = "attachments"
os.makedirs(ATTACH_DIR, exist_ok=True)

TARGET_FILE = "IT 2025.xlsx"
TARGET_SUBJECT_KEYWORD = "IT Shift Schedule"

# Connect to iCloud Mail via IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select("inbox")

# Search for all emails (can also filter UNSEEN if needed)
status, messages = mail.search(None, "ALL")
msg_ids = messages[0].split()

# Store matching emails with their dates
matching_emails = []

for num in msg_ids:
    status, data = mail.fetch(num, "(RFC822)")
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    subject = msg["subject"] or ""
    date = parsedate_to_datetime(msg["date"])
    
    if TARGET_SUBJECT_KEYWORD in subject:
        matching_emails.append((date, msg))

# If any matching emails found, take the most recent
if matching_emails:
    # Sort by date descending
    matching_emails.sort(key=lambda x: x[0], reverse=True)
    most_recent_msg = matching_emails[0][1]
    print(f"Processing most recent email: {most_recent_msg['subject']}")

    for part in most_recent_msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename == TARGET_FILE:
                filepath = os.path.join(ATTACH_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {filepath}")
else:
    print("No emails found with the specified subject keyword.")

mail.close()
mail.logout()
