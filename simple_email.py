"""
Simple Email Sender

Requirements:
    - SMTP server settings
    - Email account credentials (if required by SMTP server)

Usage:
    python simple_email.py
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Email server settings
# Common SMTP servers:
# - Gmail:         smtp.gmail.com  (requires TLS and auth)
# - Outlook/Live:  smtp.office365.com (requires TLS and auth)
# - Local relay:   localhost or IP or relay.my-service.com (typically port 25, no auth)
SMTP_SERVER = "smtp.my-service.com"
# Common SMTP ports:
# - 587 - TLS (most common for authenticated SMTP)
# - 465 - SSL (legacy, not recommended)
# - 25  - Default SMTP port (commonly used by SMTP relays)
SMTP_PORT = 25
USE_TLS = False   # True for TLS, False for no encryption or SSL
USE_AUTH = False  # True if server requires authentication

# Email credentials (only needed if USE_AUTH is True)
EMAIL_USER = "sender@example.com"
EMAIL_PASSWORD = "your_password"

# Email details
FROM_EMAIL = "sender@example.com"  # Can be different from EMAIL_USER
TO_EMAIL = "recipient@example.com"
SUBJECT = "Test Email"
BODY = "This is a test email."

# Create message
msg = MIMEMultipart()
msg['From'] = FROM_EMAIL
msg['To'] = TO_EMAIL
msg['Subject'] = SUBJECT
msg.attach(MIMEText(BODY, 'plain'))

# Send email
try:
    server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
    if USE_TLS:
        server.starttls()
    if USE_AUTH:
        server.login(EMAIL_USER, EMAIL_PASSWORD)
    server.send_message(msg)
    server.quit()
    print("Email sent successfully!")
except Exception as e:
    print(f"Error sending email: {str(e)}")