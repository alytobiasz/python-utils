#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Simple Gmail Sender

Requirements:
    - Gmail account with App Password enabled
    (To get an App Password: Google Account > Security > 2-Step Verification > App passwords)

Usage:
    python simple_gmail.py
"""

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Gmail credentials
GMAIL_USER = "your.email@gmail.com"
APP_PASSWORD = "your_16_char_app_password"

# Email details
TO_EMAIL = "recipient@example.com"
SUBJECT = "Test Email"
BODY = "This is a test email sent from Python."

# Create message
msg = MIMEMultipart()
msg['From'] = GMAIL_USER
msg['To'] = TO_EMAIL
msg['Subject'] = SUBJECT
msg.attach(MIMEText(BODY, 'plain'))

# Send email
try:
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(GMAIL_USER, APP_PASSWORD)
        server.send_message(msg)
    print("Email sent successfully!")
except Exception as e:
    print(f"Error sending email: {str(e)}") 