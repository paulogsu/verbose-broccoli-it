import imaplib
import email
from email import policy
import os
import re
import sys
import datetime
import time
import glob
import pandas as pd
from icalendar import Calendar, Event
from datetime import datetime, timedelta

# -------------------- CONFIGURATION --------------------
EMAIL_USER = os.environ.get("EMAIL_USER")
EMAIL_PASS = os.environ.get("EMAIL_PASS")
IMAP_SERVER = os.environ.get("IMAP_SERVER", "imap.gmail.com")
IMAP_PORT = int(os.environ.get("IMAP_PORT", 993))
TARGET_SENDER = os.environ.get("TARGET_SENDER", "paulo.costa@dlh.de")

ATTACHMENTS_DIR = os.path.join(os.getcwd(), "attachments")
os.makedirs(ATTACHMENTS_DIR, exist_ok=True)

OUTPUT_ICAL_FILE = "it.ics"
SCHEDULE_YEAR = 2025
TEAM_MEMBERS = [
    "LUISA TAVARES","FABIO FILIPE PEREIRA","ARNOLD LUANZAMBI",
    "BEATRIZ SEIXAS","PAULO COSTA","PEDRO SOARES",
    "JOSE ALEXANDRE FERREIRA","RICHARD JESUS"
]

SHIFT_TIMES = {'P06':'06:00-14:00','P08':'08:00-16:00','P09':'09:00-17:00','P14':'14:00-22:00','P22':'22:00-06:00'}
MONTH_MAP = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,'sep':9,'oct':10,'nov':11,'dec':12}

# -------------------- HELPER FUNCTIONS --------------------
def sanitize_filename(filename):
    return re.sub(r"[^a-zA-Z0-9._-]", "_", filename)

def is_shift(code):
    code = str(code).strip().upper()
    return code.startswith('P') and any(c.isdigit() for c in code[1:]) if len(code) >= 3 else False

def get_month(sheet):
    return MONTH_MAP.get(sheet[:3].lower())

def shift_desc(code):
    return f"{code} ({SHIFT_TIMES[code]})" if code in SHIFT_TIMES else code

def process_sheet(df, month, person):
    events = []
    person_row = df[df.iloc[:,0].str.strip().str.upper() == person.upper()]
    if person_row.empty: return events
    person_row = person_row.iloc[0]
    day_row = df[df.iloc[:,0].apply(lambda x: any(c.isdigit() for c in str(x)))].iloc[0]
    for i in range(1, len(person_row)):
        try:
            day = int(day_row.iloc[i])
            code = str(person_row.iloc[i]).strip()
            if code and is_shift(code):
                events.append({'day':day,'month':month,'person':person,'code':code,'desc':shift_desc(code)})
        except: continue
    return events

# -------------------- FETCH EMAIL ATTACHMENTS --------------------
if not all([EMAIL_USER, EMAIL_PASS, IMAP_SERVER, TARGET_SENDER]):
    sys.exit("ERROR: EMAIL_USER, EMAIL_PASS, IMAP_SERVER, and TARGET_SENDER must be set.")

print("Connecting to IMAP server...")
try:
    with imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT) as mail:
        mail.login(EMAIL_USER, EMAIL_PASS)
        mail.select("inbox")

        status, messages = mail.search(None, f'(FROM "{TARGET_SENDER}")')
        if status != "OK":
            sys.exit("Failed to search emails.")

        msg_ids = messages[0].split()
        if not msg_ids:
            print(f"No messages found from {TARGET_SENDER}.")
            sys.exit(0)

        newest_msg_id = msg_ids[-1]
        print(f"Processing newest message ID: {newest_msg_id.decode()} from {TARGET_SENDER}")

        status, data = mail.fetch(newest_msg_id, "(RFC822)")
        if status != "OK" or not data:
            sys.exit("Failed to fetch the newest message.")

        raw_email = None
        for item in data:
            if isinstance(item, tuple) and item[1]:
                raw_email = item[1]
                break

        if not raw_email:
            sys.exit("No email bytes found for the newest message.")

        msg = email.message_from_bytes(raw_email, policy=policy.default)
        subject = msg.get("subject", "(no subject)")
        print(f"Subject: {subject}")

        found_any = False
        for part in msg.walk():
            filename = part.get_filename()
            if filename:
                filename = sanitize_filename(filename)
                filepath = os.path.join(ATTACHMENTS_DIR, filename)
                with open(filepath, "wb") as f:
                    f.write(part.get_payload(decode=True))
                    f.flush()
                    os.fsync(f.fileno())  # ensure fully written
                print(f"Saved attachment: {filepath}")

                # Pause to ensure file is fully written
                print("Pausing 30 seconds to ensure file write completion...")
                time.sleep(30)

                found_any = True

        if not found_any:
            print("No attachments found in the newest message.")

        mail.logout()

except imaplib.IMAP4.error as e:
    sys.exit(f"IMAP error: {e}")

# -------------------- BUILD CALENDAR --------------------
# Look for the latest IT_2025 Excel file in attachments/
excel_files = sorted(glob.glob(os.path.join(ATTACHMENTS_DIR, "IT_2025*.xlsx")), reverse=True)
if not excel_files:
    sys.exit(f"ERROR: No IT_2025.xlsx file found in {ATTACHMENTS_DIR}/")
INPUT_EXCEL_FILE = excel_files[0]
print(f"Using Excel file: {INPUT_EXCEL_FILE}")

try:
    xls = pd.ExcelFile(INPUT_EXCEL_FILE)
except Exception as e:
    sys.exit(f"ERROR: Cannot read Excel file {INPUT_EXCEL_FILE}: {e}")

cal = Calendar()
cal.add('prodid','-//IT Team Schedule//mxm.dk//')
cal.add('version','2.0')
all_events = []

for sheet in xls.sheet_names:
    month = get_month(sheet)
    if not month: continue
    try:
        df = pd.read_excel(INPUT_EXCEL_FILE, sheet_name=sheet)
    except: continue
    for person in TEAM_MEMBERS:
        all_events.extend(process_sheet(df, month, person))

for e in all_events:
    try:
        event = Event()
        start = datetime(SCHEDULE_YEAR, e['month'], e['day'])
        event.add("summary", f"{e['code']} - {e['person']}")
        event.add("dtstart", start.date())
        event.add("dtend", (start + timedelta(days=1)).date())
        event.add("description", f"{e['person']} working {e['desc']}")
        cal.add_component(event)
    except: continue

if all_events:
    with open(OUTPUT_ICAL_FILE, "wb") as f:
        f.write(cal.to_ical())
    print(f"Calendar created with {len(all_events)} events: {OUTPUT_ICAL_FILE}")
else:
    print("No events created. Check Excel data, sheet names, and team member names.")
