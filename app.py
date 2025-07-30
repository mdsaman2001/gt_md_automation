import imaplib
import email
import os
import re
import zipfile
import csv
from PyPDF2 import PdfReader
from datetime import datetime

# --- Configuration ---
EMAIL = "sanket26102000@gmail.com"
PASSWORD = "rzafeqjbofimosoq"  # Use App Password if 2FA is enabled
IMAP_SERVER = "imap.gmail.com"
DOWNLOAD_FOLDER = "downloads"

# --- Ensure download directory exists ---
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

# --- Connect to Gmail ---
mail = imaplib.IMAP4_SSL(IMAP_SERVER)
mail.login(EMAIL, PASSWORD)
mail.select("inbox")

# --- Search for unread emails ---
status, messages = mail.search(None, '(UNSEEN)')
email_ids = messages[0].split()


# --- Collect all emails ---
all_emails = set()

# --- Process each unread email ---
for e_id in email_ids:
    status, msg_data = mail.fetch(e_id, '(RFC822)')
    raw_email = msg_data[0][1]
    msg = email.message_from_bytes(raw_email)

    for part in msg.walk():
        if part.get_content_maintype() == 'multipart':
            continue
        if part.get('Content-Disposition') is None:
            continue

        filename = part.get_filename()
        if filename and filename.endswith('.zip'):
            filepath = os.path.join(DOWNLOAD_FOLDER, filename)

            # --- Save the ZIP file ---
            with open(filepath, 'wb') as f:
                f.write(part.get_payload(decode=True))
            print(f"Downloaded ZIP: {filepath}")

            # --- Extract ZIP contents ---
            with zipfile.ZipFile(filepath, 'r') as zip_ref:
                zip_ref.extractall(DOWNLOAD_FOLDER)
            print(f"Extracted ZIP: {filepath}")

            # --- Find and process PDF ---
            for extracted_file in zip_ref.namelist():
                if extracted_file.endswith('.pdf'):
                    pdf_path = os.path.join(DOWNLOAD_FOLDER, extracted_file)

                    # --- Extract email addresses from PDF ---
                    with open(pdf_path, 'rb') as pdf_file:
                        reader = PdfReader(pdf_file)
                        text = ""
                        for page in reader.pages:
                            text += page.extract_text() or ""

                        # --- Extract email addresses using regex ---
                        email_pattern = r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}"
                        found_emails = re.findall(email_pattern, text)
                        print(f"📧 Emails found in {extracted_file}:")
                        for email_id in found_emails:
                            print(email_id)
                            all_emails.add(email_id)
            print("------")

    # Mark the email as seen after extraction is done
    mail.store(e_id, '+FLAGS', '\\Seen')

# --- Write all emails to CSV in extract folder with timestamp ---
if all_emails:
    extract_folder = 'extract'
    os.makedirs(extract_folder, exist_ok=True)
    date_str = datetime.now().strftime('%d-%m-%Y')
    csv_filename = f"extracted_emails_{date_str}.csv"
    csv_path = os.path.join(extract_folder, csv_filename)
    with open(csv_path, 'w', newline='', encoding='utf-8') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(['Email'])
        for email_id in sorted(all_emails):
            writer.writerow([email_id])
    print(f"Saved {len(all_emails)} emails to {csv_path}")

# --- Logout ---
mail.logout()
