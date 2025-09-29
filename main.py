import os
import re
import sys
from pyicloud_ipd import PyiCloudService

# Load iCloud credentials from environment variables
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")

if not EMAIL_USER or not EMAIL_PASS:
    sys.exit("ERROR: EMAIL_USER and EMAIL_PASS must be set as secrets.")

# Folder to save attachments
ATTACH_DIR = os.path.join(os.getcwd(), "attachments")
os.makedirs(ATTACH_DIR, exist_ok=True)

def sanitize_filename(filename):
    """Sanitize filenames to prevent invalid characters."""
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

print("Logging into iCloud...")
api = PyiCloudService(EMAIL_USER, EMAIL_PASS)

# If 2FA is enabled, the first run will request code via console
if api.requires_2fa:
    print("Two-factor authentication required. Please provide the code sent to your device.")
    code = input("Enter 2FA code: ")
    result = api.validate_2fa_code(code)
    if not result:
        sys.exit("Failed 2FA authentication.")
    api.trust_session()

print("Fetching all inbox messages...")
# pyicloud-ipd provides mail access via api.mail
# We iterate all mail folders, here focusing on inbox
inbox = api.mail.inbox

found_any = False

for message in inbox.all():
    sender = message.sender
    subject = message.subject
    print(f"Processing message from {sender}, subject: {subject}")

    for attachment in message.attachments:
        filename = sanitize_filename(attachment.filename)
        filepath = os.path.join(ATTACH_DIR, filename)

        with open(filepath, "wb") as f:
            f.write(attachment.file.read())
        print(f"Saved attachment: {filepath}")
        found_any = True

if not found_any:
    print("No attachments found in the inbox.")
