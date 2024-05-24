import pandas as pd
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

def send_email_with_attachment(sheet_path, recipient_email):
    # Sender email credentials
    sender_email = 'iconcreationai81@gmail.com'
    password = 'tetkkfrshnidygvq'

    # Create the email headers and body
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = 'Your Requested Sheet Attachment'

    # Email body
    body = 'Please find the attached sheet as requested.'
    message.attach(MIMEText(body, 'plain'))

    # Attach the sheet
    attachment = open(sheet_path, 'rb')
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename={os.path.basename(sheet_path)}')
    message.attach(part)
    attachment.close()

    # Sending the email
    try:
        session = smtplib.SMTP('smtp.gmail.com', 587)
        session.starttls()
        session.login(sender_email, password)
        text = message.as_string()
        session.sendmail(sender_email, recipient_email, text)
        session.quit()
        print('Mail Sent Successfully')
    except Exception as e:
        print(f'Failed to send email: {e}')

# Example usage
sheet_path = 'path_to_your_excel_sheet.xlsx'  # Replace with the path to your Excel sheet
recipient_email = 'recipient@example.com'  # Replace with the recipient's email address
send_email_with_attachment(sheet_path, recipient_email)
