import imaplib
import email
import os

EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.mail.me.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))

ATTACH_DIR = "attachments"
os.makedirs(ATTACH_DIR, exist_ok=True)

mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
mail.login(EMAIL_USER, EMAIL_PASS)
mail.select("inbox")

# Search for recent unseen emails (adjust the search as needed)
status, messages = mail.search(None, '(UNSEEN)')
msg_ids = messages[0].split()

for num in msg_ids:
    status, data = mail.fetch(num, "(RFC822)")
    raw_email = data[0][1]
    msg = email.message_from_bytes(raw_email)

    subject = msg["subject"]
    print("Processing:", subject)

    for part in msg.walk():
        if part.get_content_disposition() == "attachment":
            filename = part.get_filename()
            if filename:
                filepath = os.path.join(ATTACH_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                print(f"Saved attachment: {filepath}")

mail.close()
mail.logout()
