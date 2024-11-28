import imaplib
import email
from email.header import decode_header
from openpyxl import Workbook
from dotenv import load_dotenv
import os

# Load environment variables from .env file
load_dotenv()

EMAIL = os.getenv('EMAIL')
PASSWORD = os.getenv('PASSWORD')
IMAP_SERVER = 'imap.gmail.com'

# Function to decode the subject
def decode_subject(subject):
    decoded_fragments = decode_header(subject)
    subject_string = ""
    for fragment, encoding in decoded_fragments:
        if isinstance(fragment, bytes):
            # Handle unknown-8bit as latin1 fallback
            if encoding and encoding.lower() == "unknown-8bit":
                subject_string += fragment.decode('latin1', errors='ignore')
            else:
                subject_string += fragment.decode(encoding if encoding else 'utf-8', errors='ignore')
        else:
            subject_string += fragment
    return subject_string


# Function to fetch emails and populate Excel
def fetch_emails_to_excel():
    # Connect to Gmail server
    mail = imaplib.IMAP4_SSL(IMAP_SERVER)
    mail.login(EMAIL, PASSWORD)
    mail.select("inbox")  # Select inbox (you can choose other folders)

    # Search for all emails
    status, email_ids = mail.search(None, "ALL")
    if status != 'OK':
        print("Failed to retrieve emails.")
        return

    email_ids = email_ids[0].split()
    wb = Workbook()
    ws = wb.active
    ws.title = "Emails"

    # Add headers
    ws.append(["Subject", "Body HTML"])

    # Fetch and process each email
    for email_id in email_ids:
        status, data = mail.fetch(email_id, "(RFC822)")
        if status != 'OK':
            print(f"Failed to fetch email with ID: {email_id}")
            continue

        raw_email = data[0][1]
        msg = email.message_from_bytes(raw_email)

        # Extract subject
        subject = decode_subject(msg["subject"])

        # Extract body HTML
        body_html = ""
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_type() == "text/html":
                    body_html = part.get_payload(decode=True).decode('utf-8', errors='ignore')
                    break
        else:
            if msg.get_content_type() == "text/html":
                body_html = msg.get_payload(decode=True).decode('utf-8', errors='ignore')

        # Append data to Excel
        ws.append([subject, body_html])

    # Save the workbook
    wb.save("emails.xlsx")
    print("Emails saved to emails.xlsx.")

    # Close the connection
    mail.logout()

if __name__ == "__main__":
    fetch_emails_to_excel()
