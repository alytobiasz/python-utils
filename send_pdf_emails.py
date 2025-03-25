"""
PDF Email Sender

This script sends emails with PDF attachments to recipients based on a CSV mapping file.
The CSV file should have two columns:
    - email_column: Contains recipient email addresses
    - pdf_column: Contains the corresponding PDF filenames

Requirements:
    - Python 3.6+
    - Valid SMTP server credentials

Usage:
    python send_pdf_emails.py <config_file>

Example config.txt:
    smtp_server = smtp.gmail.com
    smtp_port = 587
    smtp_username = your.email@gmail.com
    smtp_password = your_password
    email_subject = Your Document
    email_body = Please find your document attached.
    input_directory = path/to/pdf/files
    mapping_file = path/to/mapping.csv
    email_column = Email Address
    pdf_column = PDF Filename
    test_mode = true  # Optional - if true, prints email info without sending

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

def read_config(config_path):
    """Read and validate the configuration file."""
    config = {}
    required_fields = [
        'smtp_server', 'smtp_port', 'smtp_username', 'smtp_password',
        'email_subject', 'email_body', 'input_directory', 'mapping_file',
        'email_column', 'pdf_column'
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
        
        # Set default for test_mode
        config['test_mode'] = config.get('test_mode', '').lower() == 'true'
        
        return config
        
    except Exception as e:
        raise ValueError(f"Error reading config file: {str(e)}")

def read_mapping_file(mapping_file, email_column, pdf_column):
    """Read the CSV mapping file and return a dictionary of PDF filenames to email addresses."""
    mapping = {}
    try:
        with open(mapping_file, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            
            # Verify required columns exist
            if email_column not in reader.fieldnames:
                raise ValueError(f"Email column '{email_column}' not found in mapping file")
            if pdf_column not in reader.fieldnames:
                raise ValueError(f"PDF column '{pdf_column}' not found in mapping file")
            
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
                        
                    mapping[pdf_file] = email
        
        return mapping
        
    except Exception as e:
        raise ValueError(f"Error reading mapping file: {str(e)}")

def send_email(smtp_config, to_email, subject, body, attachment_path, test_mode=False):
    """Send an email with a PDF attachment."""
    # Create message
    msg = MIMEMultipart()
    msg['From'] = smtp_config['smtp_username']
    msg['To'] = to_email
    msg['Subject'] = subject
    
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
    
    if test_mode:
        print(f"\nWould send email:")
        print(f"To: {to_email}")
        print(f"Subject: {subject}")
        print(f"Attachment: {filename}")
        return True
    
    # Send email
    try:
        with smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port']) as server:
            server.starttls()
            server.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
            server.send_message(msg)
        return True
    except Exception as e:
        print(f"Error sending email to {to_email}: {str(e)}")
        return False

def main():
    if len(sys.argv) != 2:
        print("Usage: python send_pdf_emails.py <config_file>")
        sys.exit(1)
    
    try:
        start_time = time.time()
        
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
        
        # Process PDF files
        success_count = 0
        total_count = 0
        skipped_count = 0
        not_found_count = 0
        
        print(f"\nProcessing PDF files in: {input_dir}")
        if config['test_mode']:
            print("Running in TEST MODE - no emails will be sent")
        
        # Get list of available PDF files
        available_pdfs = {f.lower(): f for f in os.listdir(input_dir) if f.lower().endswith('.pdf')}
        
        # Process each mapping
        for pdf_file, email in mapping.items():
            total_count += 1
            pdf_lower = pdf_file.lower()
            
            # Check if PDF exists (case-insensitive)
            if pdf_lower not in available_pdfs:
                print(f"\nWarning: PDF file not found: {pdf_file}")
                not_found_count += 1
                continue
            
            actual_filename = available_pdfs[pdf_lower]
            print(f"\nProcessing: {actual_filename} -> {email}")
            file_path = os.path.join(input_dir, actual_filename)
            
            # Send email
            if send_email(
                smtp_config={
                    'smtp_server': config['smtp_server'],
                    'smtp_port': config['smtp_port'],
                    'smtp_username': config['smtp_username'],
                    'smtp_password': config['smtp_password']
                },
                to_email=email,
                subject=config['email_subject'],
                body=config['email_body'],
                attachment_path=file_path,
                test_mode=config['test_mode']
            ):
                success_count += 1
        
        # Print completion message
        elapsed_time = time.time() - start_time
        print(f"\nProcessing completed in {elapsed_time:.1f} seconds")
        print(f"Emails {('would be ' if config['test_mode'] else '')}sent: {success_count}/{total_count}")
        if not_found_count > 0:
            print(f"PDF files not found: {not_found_count}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 