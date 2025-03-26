"""
PDF Email Sender

This script sends emails with PDF attachments to recipients based on a CSV mapping file.
The CSV file should have two columns:
    - email_column: Contains recipient email addresses
    - pdf_column: Contains the corresponding PDF filenames

Requirements:
    - Python 3.6+
    - SMTP server settings

Usage:
    python send_pdf_emails.py <config_file>

Example config.txt:
    # SMTP Server Settings
    smtp_server = smtp.my-relay.com     # e.g., smtp.gmail.com for Gmail, smtp.office365.com for Outlook
    smtp_port = 25                      # 25 for non-TLS, 587 for TLS (Gmail/Outlook)
    use_tls = false                     # true for Gmail/Outlook, false for basic SMTP relay
    use_auth = false                    # true if server requires username/password
    
    # Authentication (required only if use_auth = true)
    smtp_username = your.email@gmail.com
    smtp_password = your_app_password
    
    # Email Settings
    from_email = sender@my-domain.com   # Email address to send from (required)
    email_subject = Your Document
    email_body_file = email_body.txt    # Path to file containing email body text
    
    # File Locations
    input_directory = path/to/pdf/files
    mapping_file = path/to/mapping.csv
    email_column = Email Address
    pdf_column = PDF Filename
    
    # Optional Settings
    bcc_recipients = archive@company.com, supervisor@company.com  # Optional - comma-separated list
    test_mode = true                    # Optional - if true, prints email info without sending

Note for Gmail Users:
    If using Gmail, you must use an App Password instead of your regular password.
    To generate an App Password:
    1. Go to your Google Account settings
    2. Navigate to Security > 2-Step Verification
    3. Scroll to the bottom and click on "App passwords"
    4. Select "Mail" and your device
    5. Copy the generated 16-character password
    6. Use this password in the config file
"""

import sys
import os
import smtplib
import time
import csv
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

def read_email_body(body_file):
    """Read the email body text from a file."""
    try:
        with open(body_file, 'r', encoding='utf-8') as f:
            return f.read().strip()
    except Exception as e:
        raise ValueError(f"Error reading email body file: {str(e)}")

def read_config(config_path):
    """Read and validate the configuration file."""
    config = {}
    required_fields = [
        'smtp_server', 'smtp_port', 'input_directory', 'mapping_file',
        'email_column', 'pdf_column', 'email_subject', 'email_body_file',
        'from_email'  # Required for From address when not using authentication
    ]
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    key, value = [x.strip() for x in line.split('=', 1)]
                    config[key] = value
        
        # Validate required fields
        missing = [field for field in required_fields if field not in config]
        if missing:
            raise ValueError(f"Missing required fields in config: {', '.join(missing)}")
        
        # Convert port to integer
        config['smtp_port'] = int(config['smtp_port'])
        
        # Set defaults for optional settings
        config['test_mode'] = config.get('test_mode', '').lower() == 'true'
        config['use_tls'] = config.get('use_tls', '').lower() == 'true'
        config['use_auth'] = config.get('use_auth', '').lower() == 'true'
        
        # Validate auth credentials if auth is enabled
        if config['use_auth']:
            if 'smtp_username' not in config or 'smtp_password' not in config:
                raise ValueError("SMTP username and password are required when use_auth is true")
        
        # Read email body from file
        config['email_body'] = read_email_body(config['email_body_file'])
        
        # Handle BCC recipients
        if 'bcc_recipients' in config:
            # Split by comma and strip whitespace
            config['bcc_recipients'] = [
                email.strip() 
                for email in config['bcc_recipients'].split(',')
                if email.strip()
            ]
        else:
            config['bcc_recipients'] = []
        
        return config
        
    except Exception as e:
        raise ValueError(f"Error reading config file: {str(e)}")

def read_mapping_file(mapping_file, email_column, pdf_column):
    """Read the CSV mapping file and return a dictionary of PDF filenames to lists of email addresses."""
    mapping = {}
    try:
        with open(mapping_file, 'r', encoding='utf-8-sig') as f:  # Changed to utf-8-sig to handle BOM
            reader = csv.DictReader(f)
            
            # Debug print
            print(f"Looking for columns: '{email_column}' and '{pdf_column}'")
            print(f"\nFound CSV columns: {reader.fieldnames}")
            
            # Clean up fieldnames to remove any BOM characters
            reader.fieldnames = [field.strip('\ufeff') for field in reader.fieldnames]
            
            # Verify required columns exist
            if email_column not in reader.fieldnames:
                raise ValueError(f"Email column '{email_column}' not found in mapping file. Available columns: {reader.fieldnames}")
            if pdf_column not in reader.fieldnames:
                raise ValueError(f"PDF column '{pdf_column}' not found in mapping file. Available columns: {reader.fieldnames}")
            
            # Read mappings
            for row in reader:
                email = row[email_column]
                pdf_file = row[pdf_column]
                
                if email and pdf_file:
                    email = str(email).strip()
                    pdf_file = str(pdf_file).strip()
                    
                    if '@' not in email:
                        print(f"Warning: Skipping invalid email address: {email}")
                        continue
                    
                    # Initialize list for PDF if not exists
                    if pdf_file not in mapping:
                        mapping[pdf_file] = []
                    
                    # Add email to list if not already present
                    if email not in mapping[pdf_file]:
                        mapping[pdf_file].append(email)
        
        return mapping
        
    except Exception as e:
        raise ValueError(f"Error reading mapping file: {str(e)}")

def send_email(smtp_config, to_email, subject, body, attachment_path, test_mode=False, progress=None):
    """Send an email with a PDF attachment."""
    start_time = time.time()
    
    # Create message
    msg = MIMEMultipart()
    # Use smtp_username if auth is enabled, otherwise use from_email from config
    msg['From'] = smtp_config.get('smtp_username') if smtp_config.get('use_auth') else smtp_config['from_email']
    msg['To'] = to_email
    msg['Subject'] = subject
    
    # Add BCC recipients if configured
    if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
        msg['Bcc'] = ', '.join(smtp_config['bcc_recipients'])
    
    # Add body
    msg.attach(MIMEText(body, 'plain'))
    
    # Add attachment
    with open(attachment_path, 'rb') as f:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(f.read())
    
    encoders.encode_base64(part)
    filename = os.path.basename(attachment_path)
    part.add_header(
        'Content-Disposition',
        f'attachment; filename= {filename}'
    )
    msg.attach(part)
    
    progress_info = f"[{progress[0]}/{progress[1]}] " if progress else ""
    
    if test_mode:
        print(f"\n{progress_info}Would send email:")
        print(f"From: {msg['From']}")
        print(f"To: {to_email}")
        if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
            print(f"Bcc: {', '.join(smtp_config['bcc_recipients'])}")
        print(f"Subject: {subject}")
        print(f"Attachment: {filename}")
        return True
    
    # Send email
    try:
        with smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port']) as server:
            if smtp_config.get('use_tls', False):
                server.starttls()
            if smtp_config.get('use_auth', False):
                server.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
            server.send_message(msg)
            
        elapsed_time = time.time() - start_time
        print(f"{progress_info}Email sent to {to_email} ({filename}) - took {elapsed_time:.2f} seconds")
        return True
    except Exception as e:
        elapsed_time = time.time() - start_time
        print(f"{progress_info}Error sending email to {to_email} ({filename}) after {elapsed_time:.2f} seconds: {str(e)}")
        return False

def main():
    if len(sys.argv) != 2:
        print("Usage: python send_pdf_emails.py <config_file>")
        sys.exit(1)
    
    try:
        start_time = time.time()
        total_email_time = 0
        
        # Read configuration
        config = read_config(sys.argv[1])
        input_dir = config['input_directory']
        
        # Verify input directory exists
        if not os.path.isdir(input_dir):
            raise ValueError(f"Input directory does not exist: {input_dir}")
        
        # Read mapping file
        print("\nReading mapping file...")
        mapping = read_mapping_file(
            config['mapping_file'],
            config['email_column'],
            config['pdf_column']
        )
        
        if not mapping:
            raise ValueError("No valid mappings found in mapping file")
        
        # Calculate total number of emails to send
        total_emails = sum(len(emails) for emails in mapping.values())
        
        # Process PDF files
        success_count = 0
        current_count = 0
        skipped_count = 0
        not_found_count = 0
        
        print(f"\nProcessing {total_emails} emails using {len(mapping)} PDF files from: {input_dir}")
        print(f"SMTP Server: {config['smtp_server']}:{config['smtp_port']}")
        print(f"From: {config.get('smtp_username') if config.get('use_auth') else config['from_email']}")
        print(f"TLS: {'Enabled' if config.get('use_tls') else 'Disabled'}")
        print(f"Authentication: {'Enabled' if config.get('use_auth') else 'Disabled'}")
        if config['test_mode']:
            print("(TEST MODE - Emails will not be sent)")
        if config['bcc_recipients']:
            print(f"BCC recipients: {', '.join(config['bcc_recipients'])}")
        
        for pdf_file, emails in mapping.items():
            pdf_path = os.path.join(input_dir, pdf_file)
            
            if not os.path.isfile(pdf_path):
                print(f"PDF file not found: {pdf_file} (skipping {len(emails)} recipient(s))")
                not_found_count += len(emails)
                continue
            
            # Send to each recipient for this PDF
            for email in emails:
                current_count += 1
                email_start_time = time.time()
                success = send_email(
                    smtp_config={
                        'smtp_server': config['smtp_server'],
                        'smtp_port': config['smtp_port'],
                        'use_tls': config.get('use_tls', False),
                        'use_auth': config.get('use_auth', False),
                        'smtp_username': config.get('smtp_username', ''),
                        'smtp_password': config.get('smtp_password', ''),
                        'from_email': config['from_email'],
                        'bcc_recipients': config['bcc_recipients']
                    },
                    to_email=email,
                    subject=config['email_subject'],
                    body=config['email_body'],
                    attachment_path=pdf_path,
                    test_mode=config['test_mode'],
                    progress=(current_count, total_emails)
                )
                
                if success:
                    success_count += 1
                    total_email_time += (time.time() - email_start_time)
                else:
                    skipped_count += 1
        
        # Print summary
        total_time = time.time() - start_time
        avg_email_time = total_email_time / success_count if success_count > 0 else 0
        
        print("\nSummary:")
        print(f"Total time: {total_time:.2f} seconds")
        print(f"Average time per email: {avg_email_time:.2f} seconds")
        print(f"Total emails to send: {total_emails}")
        print(f"Successfully sent: {success_count}")
        print(f"Files not found: {not_found_count}")
        print(f"Errors/skipped: {skipped_count}")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)

if __name__ == '__main__':
    main() 