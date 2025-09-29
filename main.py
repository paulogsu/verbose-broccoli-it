import imaplib
import email
import os

EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

ATTACH_DIR = "attachments"

# Ensure attachments folder exists
if not os.path.exists(ATTACH_DIR):
    os.makedirs(ATTACH_DIR)

TARGET_SENDER = "paulo.costa@dlh.de"

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

found_any = False

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
    sender = msg.get("from", "(unknown sender)")
    subject = msg.get("subject", "(no subject)")
    print(f"Processing email {num} from {sender}, subject: {subject}")

    # Only process emails from the target sender
    if TARGET_SENDER in sender:
        for part in msg.walk():
            # Check both standard attachments and inline files
            filename = part.get_filename()
            content_disposition = part.get("Content-Disposition", "")
            if filename:
                filepath = os.path.join(ATTACH_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {filepath}")
                found_any = True

if not found_any:
    print(f"No attachments found from sender: {TARGET_SENDER}")

mail.close()
mail.logout()
