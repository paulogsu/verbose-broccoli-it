import imaplib
import email
from email import policy
import os
import re
import sys

# Load configuration from environment variables
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
TARGET_SENDER = os.environ.get("TARGET_SENDER")

# Validate required environment variables
if not all([EMAIL_USER, EMAIL_PASS, IMAP_SERVER, TARGET_SENDER]):
    sys.exit("ERROR: EMAIL_USER, EMAIL_PASS, IMAP_SERVER, and TARGET_SENDER must be set.")

# Directory to save attachments
ATTACH_DIR = os.path.join(os.getcwd(), "attachments")
os.makedirs(ATTACH_DIR, exist_ok=True)

def sanitize_filename(filename):
    """Sanitize filenames to prevent path traversal."""
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

print("Starting script...")
print(f"Connecting to IMAP server: {IMAP_SERVER}:{IMAP_PORT}")

try:
    with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as mail:
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        status, messages = mail.search(None, "ALL")
        if status != "OK":
            sys.exit("Failed to fetch emails.")

        msg_ids = messages[0].split()
        print(f"Found {len(msg_ids)} emails in the inbox.")

        found_any = False

        # Process emails from newest to oldest
        for num in reversed(msg_ids):
            status, data = mail.fetch(num, "(RFC822)")
            if status != "OK":
                print(f"Failed to fetch email {num}, skipping.")
                continue

            raw_email = next((part[1] for part in data if isinstance(part, tuple) and isinstance(part[1], bytes)), None)
            if not raw_email:
                print(f"Skipping email {num}: no email bytes found")
                continue

            msg = email.message_from_bytes(raw_email, policy=policy.default)
            sender = msg.get("from", "(unknown sender)")
            subject = msg.get("subject", "(no subject)")
            print(f"Processing email {num} from {sender}, subject: {subject}")

            # Only process emails from the target sender
            if TARGET_SENDER.lower() in sender.lower():
                for part in msg.walk():
                    filename = part.get_filename()
                    if filename and filename.lower().endswith(".xlsx"):
                        filename = sanitize_filename(filename)
                        filepath = os.path.join(ATTACH_DIR, filename)
                        with open(filepath, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        print(f"Saved .xlsx attachment: {filepath}")
                        found_any = True

        if not found_any:
            print(f"No .xlsx attachments found from sender: {TARGET_SENDER}")

except imaplib.IMAP4.error as e:
    sys.exit(f"IMAP error: {e}")
