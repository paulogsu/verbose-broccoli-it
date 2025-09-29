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

status, messages = mail.search(None, "ALL")
msg_ids = messages[0].split()

matching_emails = []

for num in msg_ids:
    status, data = mail.fetch(num, "(RFC822)")
    
    # data can be a list of tuples, sometimes with None/ints; extract bytes safely
    raw_email = None
    for part in data:
        if isinstance(part, tuple) and isinstance(part[1], bytes):
            raw_email = part[1]
            break

    if raw_email is None:
        continue  # skip this email if no proper bytes found

    msg = email.message_from_bytes(raw_email)
    subject = msg.get("subject", "")
    date = parsedate_to_datetime(msg.get("date"))

    if TARGET_SUBJECT_KEYWORD in subject:
        matching_emails.append((date, msg))

if matching_emails:
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
