"""
Send Emails with Attachments

This script sends emails with file attachments to recipients based on a CSV mapping file.
The CSV file should have:
   - email_column: Contains recipient email addresses
   - One or more attachment columns: Columns containing filenames to attach

Each line in the CSV file will result in one email being sent. The same email address
can appear on multiple lines, which will result in multiple emails being sent to that address.

Example CSV format:
   Email Address,Attachment 1,Attachment 2,Attachment 3
   user@example.com,document1.pdf,image.jpg,spreadsheet.xlsx
   user@example.com,document4.pdf,contract.docx,presentation.pptx

Requirements:
    - Python 3.6+
    - SMTP server settings

Usage:
    python send_emails_with_attachments.py <config_file>

Example config.txt:
    # SMTP Server Settings
    smtp_server = smtp.gmail.com        # e.g., smtp.gmail.com for Gmail, smtp.office365.com for Outlook
    smtp_port = 587                     # 25 for non-TLS, 587 for TLS (Gmail/Outlook)
    use_tls = true                      # true for Gmail/Outlook, false for basic SMTP relay
    use_auth = true                     # true if server requires username/password
    
    # Authentication (required only if use_auth = true)
    smtp_username = your.email@gmail.com
    smtp_password = your_app_password
    
    # Email Settings
    from_email = sender@my-domain.com   # Email address to send from (required)
    email_subject = Your Document
    email_body_file = email_body.txt    # Path to file containing email body text
    
    # File Locations
    input_directory = path/to/attachment/files
    mapping_file = path/to/mapping.csv
    email_column = Email Address
    
    # Attachment Column Configuration
    attachment_columns = Attachment 1, Attachment 2, Attachment 3  # Comma-separated list of column names containing filenames
    
    # Performance Settings
    max_threads = 4                     # Optional - number of concurrent email sending threads (default: 4)
    
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
import threading
import concurrent.futures
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path

# ANSI color codes for terminal output
YELLOW = "\033[93m"  # Yellow for warnings
RED = "\033[91m"     # Red for errors
GREEN = "\033[92m"   # Green for success
RESET = "\033[0m"    # Reset to default color

# Thread-local storage for SMTP connections
thread_local = threading.local()

def get_smtp_connection(smtp_config):
    """Get or create an SMTP connection for the current thread."""
    if not hasattr(thread_local, "smtp"):
        # Create new SMTP connection for this thread
        smtp = smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port'])
        if smtp_config.get('use_tls', False):
            smtp.starttls()
        if smtp_config.get('use_auth', False):
            smtp.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
        thread_local.smtp = smtp
    return thread_local.smtp

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
        'email_column', 'email_subject', 'email_body_file', 'attachment_columns',
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
        
        # Parse attachment columns
        config['attachment_columns'] = [col.strip() for col in config['attachment_columns'].split(',')]
        
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

def read_mapping_file(mapping_file, email_column, attachment_columns):
    """Read the CSV mapping file and return a list of (email, attachments) tuples for sending.
    
    Args:
        mapping_file (str): Path to the CSV mapping file
        email_column (str): Name of the column containing email addresses
        attachment_columns (list): List of column names containing filenames to attach
    
    Returns:
        list: List of (email, attachments) tuples for sending
    """
    email_tasks = []
    try:
        with open(mapping_file, 'r', encoding='utf-8-sig') as f:  # Changed to utf-8-sig to handle BOM
            reader = csv.DictReader(f)
            
            # Clean up fieldnames to remove any BOM characters
            reader.fieldnames = [field.strip('\ufeff') for field in reader.fieldnames]
            
            # Verify email column exists
            if email_column not in reader.fieldnames:
                raise ValueError(f"Email column '{email_column}' not found in mapping file. Available columns: {reader.fieldnames}")
            
            print(f"Looking for email column '{email_column}' and attachment columns: {', '.join(attachment_columns)}")
            
            # Verify all attachment columns exist
            missing_columns = [col for col in attachment_columns if col not in reader.fieldnames]
            if missing_columns:
                raise ValueError(f"Attachment column(s) not found in mapping file: {', '.join(missing_columns)}. Available columns: {reader.fieldnames}")
            
            # Read mappings - each row becomes one email task
            row_count = 0
            for row in reader:
                row_count += 1
                email = row[email_column]
                
                if not email or '@' not in email:
                    print(f"{YELLOW}WARNING: Skipping invalid email address in row {row_count}: {email}{RESET}")
                    continue
                
                email = str(email).strip()
                
                # Collect attachment files for this row
                files_for_row = []
                for attachment_column in attachment_columns:
                    filename = row[attachment_column]
                    if filename:
                        filename = str(filename).strip()
                        if filename:
                            files_for_row.append(filename)
                
                # Add to email tasks even if no attachments were specified 
                email_tasks.append((email, files_for_row))
            
            print(f"\nFound CSV columns: {reader.fieldnames}")
            print(f"Found {len(email_tasks)} email tasks to process")
            return email_tasks
        
    except Exception as e:
        raise ValueError(f"Error reading mapping file: {str(e)}")

def send_email(smtp_config, to_email, subject, body, attachment_paths, test_mode=False, progress=None):
    """Send an email with file attachments."""
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
    
    # Add attachments
    attachment_names = []
    for attachment_path in attachment_paths:
        try:
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
            
            encoders.encode_base64(part)
            filename = os.path.basename(attachment_path)
            attachment_names.append(filename)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename= {filename}'
            )
            msg.attach(part)
        except Exception as e:
            print(f"{RED}WARNING: Error attaching file {attachment_path}: {e}{RESET}")
            # Continue with other attachments
    
    progress_info = f"[{progress[0]}/{progress[1]}] " if progress else ""
    attachment_str = ", ".join(attachment_names)
    
    if test_mode:
        print(f"\n{progress_info}Would send email:")
        print(f"From: {msg['From']}")
        print(f"To: {to_email}")
        if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
            print(f"Bcc: {', '.join(smtp_config['bcc_recipients'])}")
        print(f"Subject: {subject}")
        print(f"Attachments: {attachment_str if attachment_str else 'None'}")
        return True
    
    # Send email
    try:
        if not test_mode:
            smtp = get_smtp_connection(smtp_config)
            smtp.send_message(msg)
        
        elapsed_time = time.time() - start_time
        print(f"{progress_info}Email sent to {to_email} with {len(attachment_names)} attachment(s) - took {elapsed_time:.2f} seconds")
        return True
    except Exception as e:
        elapsed_time = time.time() - start_time
        print(f"{RED}{progress_info}Error sending email to {to_email} after {elapsed_time:.2f} seconds: {str(e)}{RESET}")
        return False

def process_email(args):
    """Process a single email send operation."""
    smtp_config, email, attachment_paths, config, current_count, total_emails = args
    email_start_time = time.time()
    success = send_email(
        smtp_config=smtp_config,
        to_email=email,
        subject=config['email_subject'],
        body=config['email_body'],
        attachment_paths=attachment_paths,
        test_mode=config['test_mode'],
        progress=(current_count, total_emails)
    )
    return success, time.time() - email_start_time

def main():
    if len(sys.argv) != 2:
        print(f"{YELLOW}Usage: python send_emails_with_attachments.py <config_file>{RESET}")
        sys.exit(1)
    
    try:
        start_time = time.time()
        total_email_time = 0
        
        # Read configuration
        config = read_config(sys.argv[1])
        input_dir = config['input_directory']
        
        # Get number of threads from config or use default
        max_threads = int(config.get('max_threads', '4'))
        
        # Verify input directory exists
        if not os.path.isdir(input_dir):
            raise ValueError(f"Input directory does not exist: {input_dir}")
        
        # Read mapping file
        print("\nReading mapping file...")
        email_tasks = read_mapping_file(
            config['mapping_file'],
            config['email_column'],
            config['attachment_columns']
        )
        
        if not email_tasks:
            raise ValueError("No valid mappings found in mapping file")
        
        # Calculate statistics
        total_emails = len(email_tasks)
        total_attachments = sum(len(files) for _, files in email_tasks)
        
        # Process attachment files
        success_count = 0
        skipped_count = 0
        not_found_count = 0
        found_files = {}  # Track files that exist so we don't check multiple times
        
        print(f"\nProcessing {total_emails} emails with {total_attachments} total attachments")
        print(f"Using {max_threads} threads for parallel processing")
        print(f"SMTP Server: {config['smtp_server']}:{config['smtp_port']}")
        print(f"From: {config.get('smtp_username') if config.get('use_auth') else config['from_email']}")
        print(f"TLS: {'Enabled' if config.get('use_tls') else 'Disabled'}")
        print(f"Authentication: {'Enabled' if config.get('use_auth') else 'Disabled'}")
        print(f"Attachment columns: {', '.join(config['attachment_columns'])}")
        if config['test_mode']:
            print("(TEST MODE - Emails will not be sent)")
        if config['bcc_recipients']:
            print(f"BCC recipients: {', '.join(config['bcc_recipients'])}")
        
        # Prepare email tasks for processing
        processing_tasks = []
        current_count = 0
        
        # Process each email task
        for email, attachment_files in email_tasks:
            current_count += 1
            
            # If no attachment files specified, send email without attachments
            if not attachment_files:
                # Create the SMTP config for this task
                smtp_config = {
                    'smtp_server': config['smtp_server'],
                    'smtp_port': config['smtp_port'],
                    'use_tls': config.get('use_tls', False),
                    'use_auth': config.get('use_auth', False),
                    'smtp_username': config.get('smtp_username', ''),
                    'smtp_password': config.get('smtp_password', ''),
                    'from_email': config['from_email'],
                    'bcc_recipients': config['bcc_recipients']
                }
                
                # Add task with empty attachments list
                processing_tasks.append((smtp_config, email, [], config, current_count, total_emails))
                continue
            
            # Check which attachment files exist
            valid_file_paths = []
            invalid_files = []
            all_files_found = True
            
            for attachment_file in attachment_files:
                file_path = os.path.join(input_dir, attachment_file)
                
                # Check if we already know if this file exists
                if attachment_file in found_files:
                    if found_files[attachment_file]:  # File exists
                        valid_file_paths.append(file_path)
                    else:
                        # File is known to not exist
                        invalid_files.append(attachment_file)
                        all_files_found = False
                else:
                    # Check if file exists and store result for future use
                    if os.path.isfile(file_path):
                        valid_file_paths.append(file_path)
                        found_files[attachment_file] = True
                    else:
                        invalid_files.append(attachment_file)
                        found_files[attachment_file] = False
                        not_found_count += 1
                        all_files_found = False
            
            # Skip if any attachment files were not found
            if not all_files_found:
                print(f"{YELLOW}WARNING: Email to {email} skipped: Missing attachment file(s): {', '.join(invalid_files)}{RESET}")
                skipped_count += 1
                continue
            
            # Create the SMTP config once per task
            smtp_config = {
                'smtp_server': config['smtp_server'],
                'smtp_port': config['smtp_port'],
                'use_tls': config.get('use_tls', False),
                'use_auth': config.get('use_auth', False),
                'smtp_username': config.get('smtp_username', ''),
                'smtp_password': config.get('smtp_password', ''),
                'from_email': config['from_email'],
                'bcc_recipients': config['bcc_recipients']
            }
            
            # Add task for this email with its attachments
            processing_tasks.append((smtp_config, email, valid_file_paths, config, current_count, total_emails))
        
        # Process emails in parallel using thread pool
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            results = list(executor.map(process_email, processing_tasks))
            
            # Process results
            for success, email_time in results:
                if success:
                    success_count += 1
                    total_email_time += email_time
                else:
                    skipped_count += 1
        
        # Print summary
        total_time = time.time() - start_time
        avg_email_time = total_email_time / success_count if success_count > 0 else 0
        
        print("\nSummary:")
        print(f"Total time: {total_time:.2f} seconds")
        print(f"Average time per email: {avg_email_time:.2f} seconds")
        print(f"Total emails processed: {total_emails}")
        print(f"Total attachments: {total_attachments}")
        print(f"Successfully sent: {success_count}")
        print(f"Files not found: {not_found_count}")
        print(f"Emails skipped: {skipped_count}")
        
    except Exception as e:
        print(f"\nError: {str(e)}")
        sys.exit(1)
    finally:
        # Clean up any remaining SMTP connections
        if hasattr(thread_local, "smtp"):
            try:
                thread_local.smtp.quit()
            except:
                pass

if __name__ == '__main__':
    main() 