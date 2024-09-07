from flask import Flask, render_template, request
from flask_restful import Resource, Api
import csv
import imaplib
import email
import re

app = Flask(__name__)
api = Api(app)

def get_text_from_email(msg):
    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            if content_type == "text/plain":
                try:
                    text = part.get_payload(decode=True).decode('utf-8')
                except UnicodeDecodeError:
                    try:
                        text = part.get_payload(decode=True).decode('iso-8859-1')
                    except UnicodeDecodeError:
                        try:
                            text = part.get_payload(decode=True).decode('utf-16')
                        except UnicodeDecodeError:
                            text = part.get_payload(decode=True).decode('latin-1')
    else:
        try:
            text = msg.get_payload(decode=True).decode('utf-8')
        except UnicodeDecodeError:
            try:
                text = msg.get_payload(decode=True).decode('iso-8859-1')
            except UnicodeDecodeError:
                try:
                    text = msg.get_payload(decode=True).decode('utf-16')
                except UnicodeDecodeError:
                    text = msg.get_payload(decode=True).decode('latin-1')
    return text

def extract_data_from_text(text):
    extracted_data = {}
    name_match = re.search(r"Name:\s*([^\r\n]+)", text)
    if name_match:
        extracted_data["Name"] = name_match.group(1).strip()
    phone_match = re.search(r"Phone:\s*([\d\s-]+)", text)
    if phone_match:
        extracted_data["Phone"] = phone_match.group(1).strip()
    email_match = re.search(r"Email:\s*([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})", text)
    if email_match:
        extracted_data["Email"] = email_match.group(1).strip()
    company_match = re.search(r"Company:\s*([^\r\n]+)", text)
    if company_match:
        extracted_data["Company"] = company_match.group(1).strip()
    subject_match = re.search(r"Subject:\s*([^\r\n]+)", text)
    if subject_match:
        extracted_data["Subject"] = subject_match.group(1).strip()
    return extracted_data

class EmailExtraction(Resource):
    def post(self):
        email_address = request.form['T1']
        password = request.form['P1']
        extracted_data = []

        imap_settings = {
            "gmail.com": ("imap.gmail.com", "inbox"),
            "one.com": ("imap.one.com", "inbox"),
            "outlook.com": ("imap.outlook.com", "inbox")
        }

        domain = email_address.split('@')[-1]
        if domain in imap_settings:
            server, mailbox = imap_settings[domain]
            mail = imaplib.IMAP4_SSL(server)
            try:
                mail.login(email_address, password)
            except imaplib.IMAP4.error as e:
                return {"error": f"Failed to log in: {str(e)}"}

            mail.select(mailbox)
            status, email_ids = mail.search(None, "ALL")
            if status != 'OK':
                return {"error": "Failed to retrieve emails."}
            
            email_ids = email_ids[0].split()
            for email_id in email_ids:
                status, email_data = mail.fetch(email_id, "(RFC822)")
                if status != 'OK':
                    continue
                
                raw_email = email_data[0][1]
                msg = email.message_from_bytes(raw_email)
                text = get_text_from_email(msg)
                data = extract_data_from_text(text)
                if data:  # Only append non-empty data
                    extracted_data.append(data)
            
            mail.logout()
            
            with open("extracted_data.csv", "w", newline="", encoding="utf-8") as csv_file:
                fieldnames = ["Name", "Phone", "Email", "Company", "Subject"]
                csv_writer = csv.DictWriter(csv_file, fieldnames=fieldnames)
                csv_writer.writeheader()
                for data in extracted_data:
                    csv_writer.writerow({field: data.get(field, '') for field in fieldnames})
        
        return {"message": "Email extraction completed."}

api.add_resource(EmailExtraction, '/process-credentials')

@app.route('/')
def home():
    return render_template('index.html')

if __name__ == '__main__':
    app.run(debug=True)
