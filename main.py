import imaplib
import email
import os

EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

TARGET_FILE = "IT 2025.xlsx"
SAVE_PATH = TARGET_FILE

print("Starting script...")
print(f"Connecting to IMAP server: {IMAP_SERVER}:{IMAP_PORT}")

# Connect to iCloud Mail via IMAP
mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select("inbox")

# Search all emails in the inbox
status, messages = mail.search(None, "ALL")
msg_ids = messages[0].split()
print(f"Found {len(msg_ids)} emails in the inbox.")

found = False

# Process emails from newest to oldest
for num in reversed(msg_ids):
    status, data = mail.fetch(num, "(RFC822)")

    raw_email = None
    for part in data:
        if isinstance(part, tuple) and isinstance(part[1], bytes):
            raw_email = part[1]
            break

    if raw_email is None:
        print(f"Skipping email {num}: no email bytes found")
        continue

    msg = email.message_from_bytes(raw_email)
    subject = msg.get("subject", "(no subject)")
    print(f"Processing email {num} with subject: {subject}")

    for part in msg.walk():
        # Check both standard attachments and inline files
        content_disposition = part.get("Content-Disposition", "")
        filename = part.get_filename()
        if filename:
            print(f"Found attachment: {filename} (Content-Disposition: {content_disposition})")
            if filename == TARGET_FILE:
                with open(SAVE_PATH, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {SAVE_PATH}")
                found = True
                break
    if found:
        break  # Stop after first match

if not found:
    print(f"No attachment named '{TARGET_FILE}' found in the inbox.")

mail.close()
mail.logout()
