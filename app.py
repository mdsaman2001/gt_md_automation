import imaplib
import email
import os
import re
import zipfile
import csv
from PyPDF2 import PdfReader
from datetime import datetime

import smtplib
import openpyxl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Settings --> Actions --> General --> Workflow Permissions --> Read and Write for Actions

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

# --- Define a single regex for email extraction and validation (2 or 3 char TLD) ---
email_pattern = r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,3}\b"

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

                        # --- Extract email addresses using single regex ---
                        found_emails = re.findall(email_pattern, text)
                        print(f"📧 Emails found in {extracted_file}:")
                        for email_id in found_emails:
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
        # Collect all unique, valid emails first
        unique_emails = set()
        for email_id in all_emails:
            if (
                '..' in email_id
                or email_id.startswith(('.', '-'))
                or email_id.endswith(('.', '-'))
                or not re.match(email_pattern, email_id)
            ):
                continue
            trimmed_email = email_id.strip()
            unique_emails.add(trimmed_email.lower())
            print(f"Valid email: {trimmed_email}")
        # Write all unique emails at once
        writer.writerows([[email] for email in sorted(unique_emails)])
    print(f"Saved {len(unique_emails)} emails to {csv_path}")

# --- Logout ---
mail.logout()

# --- Send emails based on the extracted file ---
extract_folder = 'extract'
date_str = datetime.now().strftime('%d-%m-%Y')
csv_filename = f"extracted_emails_{date_str}.csv"
csv_path = os.path.join(extract_folder, csv_filename)

if os.path.exists(csv_path):
    with open(csv_path, newline='', encoding='utf-8') as csvfile:
        reader = csv.DictReader(csvfile)
        for row in reader:
            to_add = row['Email']
            fromaddr = "samansheikhwrr@gmail.com"
            from_add_password = "pzzwzeguipthpsgo"
            toaddr = to_add

            email_subject = "Application - Java | Back-End | Software Developer Role"
            resume_path = "./Mohd_Saman_Resume.pdf"
            filename = "Mohd_Saman_Resume.pdf"

            msg = MIMEMultipart('alternative')
            msg['From'] = fromaddr
            msg['To'] = toaddr
            msg['Subject'] = email_subject

            text = "Hi!\nHow are you?\nHere is the link you wanted:\nhttp://www.python.org"
            html = """
            <html>
            <body style=" font-family: Arial, Helvetica, sans-serif; font-size: 22px">
                <h4>Dear Hiring Manager,</h4>
                <h6>I hope you are doing well. <h6>
                <p style="font-weight: normal; line-height: 24px">
                    I am writing to express my interest in roles as a Java Developer, Back-End Developer, Software Developer, or Full-Stack Developer at your organization.
                    <br>
                    I am a B.Tech graduate with experience in Core Java, Spring Boot, Spring MVC, Hibernate, MySQL, and familiarity with front-end skills in HTML, CSS, JavaScript and React through internships.
                    <br>
                    I have attached my resume for your reference and would welcome the chance to discuss how I can contribute to your team.
                    <br>
                    Thank you for your time and consideration.
                </p>
                <table style="font-weight: normal; line-height: 24px;" border=1>
                    <tr>
                        <td style="font-weight: bold;">Location Preferences</td>
                        <td>Pune, Mumbai, Hyderabad, Bangalore</td>
                    </tr>
                    <tr>
                        <td style="font-weight: bold;">Notice Period</td>
                        <td> Immediate </td>
                    </tr>
                </table>
                <br>
                <p style="line-height: 24px">With Best Regards,<br></p>
                <p style="line-height: 22px; font-weight: normal;">
                    Mr. Md Saman Sheikh <br>
                    [M]: mdsamanngp@gmai.com<br>
                    [H]: +91-7219428932
                    <br>
                </p>
            </body>
            </html>
            """

            part1 = MIMEText(text, 'plain')
            part2 = MIMEText(html, 'html')
            msg.attach(part1)
            msg.attach(part2)

            try:
                with open(resume_path, "rb") as attachment:
                    p = MIMEBase('application', 'octet-stream')
                    p.set_payload(attachment.read())
                    encoders.encode_base64(p)
                    p.add_header('Content-Disposition', f"attachment; filename= %s" % filename)
                    msg.attach(p)

                s = smtplib.SMTP('smtp.gmail.com', 587)
                s.starttls()
                s.login(fromaddr, from_add_password)
                s.sendmail(fromaddr, toaddr, msg.as_string())
                s.quit()
                print(f"Sent email to {toaddr}")
            except Exception as e:
                print(f"Failed to send email to {toaddr}: {e}")
