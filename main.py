import os
import re
import sys
from imapclient import IMAPClient
import pyzmail

# Load configuration from environment variables
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

if not all([EMAIL_USER, EMAIL_PASS, IMAP_SERVER]):
    sys.exit("ERROR: EMAIL_USER, EMAIL_PASS, and IMAP_SERVER must be set.")

# Directory to save attachments
ATTACH_DIR = os.path.join(os.getcwd(), "attachments")
os.makedirs(ATTACH_DIR, exist_ok=True)

def sanitize_filename(filename):
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

print("Connecting to IMAP server...")
with IMAPClient(IMAP_SERVER, port=IMAP_PORT, ssl=True) as server:
    server.login(EMAIL_USER, EMAIL_PASS)
    server.select_folder('INBOX')

    print("Fetching all messages...")
    messages = server.search('ALL')
    print(f"Found {len(messages)} messages.")

    found_any = False

    for msgid, data in server.fetch(messages, ['RFC822']).items():
        message = pyzmail.PyzMessage.factory(data[b'RFC822'])
        sender = message.get_addresses('from')
        subject = message.get_subject()
        print(f"Processing message {msgid} from {sender}, subject: {subject}")

        # Iterate through all parts of the email to find attachments
        for part in message.mailparts:
            if part.filename:
                filename = sanitize_filename(part.filename)
                filepath = os.path.join(ATTACH_DIR, filename)
                with open(filepath, 'wb') as f:
                    f.write(part.get_payload())
                print(f"Saved attachment: {filepath}")
                found_any = True

    if not found_any:
        print("No attachments found in the inbox.")
