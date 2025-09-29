import imaplib
import email
import os

EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

TARGET_FILE = "IT 2025.xlsx"

# Save file directly in repo root
SAVE_PATH = TARGET_FILE

# Connect to iCloud Mail via IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select("inbox")

# Search all emails in the inbox
status, messages = mail.search(None, "ALL")
msg_ids = messages[0].split()

found = False

# Go through emails from newest to oldest
for num in reversed(msg_ids):
    status, data = mail.fetch(num, "(RFC822)")
    
    raw_email = None
    for part in data:
        if isinstance(part, tuple) and isinstance(part[1], bytes):
            raw_email = part[1]
            break

    if raw_email is None:
        continue

    msg = email.message_from_bytes(raw_email)

    # Walk through all parts to find attachments
    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename == TARGET_FILE:
                with open(SAVE_PATH, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {SAVE_PATH}")
                found = True
                break  # Stop after first match
    if found:
        break  # Stop searching emails after finding the file

if not found:
    print(f"No attachment named {TARGET_FILE} found in the inbox.")

mail.close()
mail.logout()
