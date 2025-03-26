"""
Simple Email Sender

Requirements:
    - Email account credentials
    - SMTP server settings from your email provider

Usage:
    python simple_email.py
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Email server settings
# Common SMTP servers:
# - Gmail:         smtp.gmail.com
# - Outlook/Live:  smtp.office365.com
SMTP_SERVER = "smtp.gmail.com"
# Common SMTP ports:
# - 587 - TLS (most common)
# - 465 - SSL
# - 25  - Default (no encryption, not recommended)  
SMTP_PORT = 587
USE_TLS = True   # True for TLS, False for no encryption or SSL

# Email credentials
EMAIL_USER = "sender@example.com"
EMAIL_PASSWORD = "your_password"

# Email details
TO_EMAIL = "recipient@example.com"
SUBJECT = "Test Email"
BODY = "This is a test email."

# Create message
msg = MIMEMultipart()
msg['From'] = EMAIL_USER
msg['To'] = TO_EMAIL
msg['Subject'] = SUBJECT
msg.attach(MIMEText(BODY, 'plain'))

# Send email
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    if USE_TLS:
        server.starttls()
    server.login(EMAIL_USER, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()
    print("Email sent successfully!")
except Exception as e:
    print(f"Error sending email: {str(e)}")
    print("\nCommon SMTP ports:")
    print("587 - TLS (most common)")
    print("465 - SSL")
    print("25  - Default (no encryption, not recommended)") 