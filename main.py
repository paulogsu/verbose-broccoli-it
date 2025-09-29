import imaplib
import email
from email import policy
import os
import re
import sys

# Load configuration from environment variables
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
    """Sanitize filenames to avoid invalid characters."""
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

print("Connecting to Gmail IMAP server...")
try:
    with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as mail:
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        # Search for emails from the target sender
        status, messages = mail.search(None, f'(FROM "{TARGET_SENDER}")')
        if status != "OK":
            sys.exit("Failed to search emails.")

        msg_ids = messages[0].split()
        print(f"Found {len(msg_ids)} messages from {TARGET_SENDER}.")

        found_any = False

        for num in reversed(msg_ids):
            status, data = mail.fetch(num, "(RFC822)")
            if status != "OK":
                print(f"Failed to fetch email {num}, skipping.")
                continue

            raw_email = None
            for item in data:
                if isinstance(item, tuple) and item[1]:
                    raw_email = item[1]
                    break

            if not raw_email:
                print(f"Skipping email {num}: no email bytes found")
                continue

            msg = email.message_from_bytes(raw_email, policy=policy.default)
            sender = msg.get("from")
            subject = msg.get("subject", "(no subject)")
            print(f"Processing email {num} from {sender}, subject: {subject}")

            # Iterate all parts of the email to find attachments
            for part in msg.walk():
                filename = part.get_filename()
                if filename:
                    filename = sanitize_filename(filename)
                    filepath = os.path.join(ATTACH_DIR, filename)
                    with open(filepath, "wb") as f:
                        f.write(part.get_payload(decode=True))
                    print(f"Saved attachment: {filepath}")
                    found_any = True

        if not found_any:
            print(f"No attachments found from sender: {TARGET_SENDER}")

        mail.logout()
except imaplib.IMAP4.error as e:
    sys.exit(f"IMAP error: {e}")
