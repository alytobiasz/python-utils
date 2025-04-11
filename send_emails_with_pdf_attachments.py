"""
Send Emails with PDF Attachments

This script sends emails with PDF file attachments to recipients based on a CSV mapping file.
The CSV file should have:
   - email_column: Contains recipient email addresses
   - One or more attachment columns: Columns containing filenames of PDF files to attach

Each line in the CSV file will result in one email being sent. The same email address
can appear on multiple lines, which will result in multiple emails being sent to that address.

Note: This script ONLY handles PDF files. All other file types will be rejected.

Example CSV format:
   Email Address,Attachment 1,Attachment 2,Attachment 3
   user@example.com,document1.pdf,report.pdf,contract.pdf
   user@example.com,document4.pdf,invoice.pdf,statement.pdf

Requirements:
    - Python 3.6+
    - SMTP server settings

Usage:
    python send_emails_with_pdf_attachments.py <config_file>

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
    attachment_columns = Attachment 1, Attachment 2, Attachment 3  # Comma-separated list of column names containing PDF filenames
    
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

import concurrent.futures
import csv
import email.charset as charset
import email.utils
import logging
import mimetypes
import os
import random
import signal
import smtplib
import socket
import sys
import threading
import time
import traceback
from datetime import datetime
from email import encoders
from email.mime.application import MIMEApplication
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from io import BytesIO, StringIO
from pathlib import Path

# Thread-local storage for SMTP connections
thread_local = threading.local()

# Maximum number of connection retries
MAX_RETRIES = 3

# Connection refresh settings - refresh connection every X emails
CONNECTION_REFRESH_COUNT = 20  # Refresh SMTP connection after this many emails

# Thread-safe exit flag
should_exit = threading.Event()

# Set up signal handler for graceful shutdown
def signal_handler(sig, frame):
    """Handle keyboard interrupt and other signals to gracefully shut down."""
    logging.info("\nKeyboard interrupt detected. Shutting down gracefully...")
    logging.info("Please wait for current tasks to finish (this may take a moment)...")
    
    # Set flag for threads to check
    should_exit.set()
    
    # Don't exit immediately, let cleanup happen in main thread
    # The thread pool will be shut down in main()

# Register signal handlers
signal.signal(signal.SIGINT, signal_handler)  # Handles Ctrl+C
signal.signal(signal.SIGTERM, signal_handler)  # Handles termination signal

# Set up logging
def setup_logging():
    """Set up logging to file and console."""
    # Create logs directory if it doesn't exist
    os.makedirs('logs', exist_ok=True)
    
    # Generate log filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    log_file = f"logs/email_sending_{timestamp}.log"
    
    # Create logger
    logger = logging.getLogger()
    logger.setLevel(logging.INFO)
    
    # Create formatter
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    
    # File handler
    file_handler = logging.FileHandler(log_file)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Console handler
    console_handler = logging.StreamHandler()
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    logging.info(f"Logging to console and file: {log_file}")
    return log_file

def create_smtp_connection(smtp_config):
    timeout = smtp_config.get('smtp_timeout', 60)
    smtp = smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port'], timeout=timeout)
    smtp.set_debuglevel(0)
    smtp.ehlo_or_helo_if_needed()
    
    if smtp_config.get('use_tls'):
        smtp.starttls()
        smtp.ehlo()

    if smtp_config.get('use_auth'):
        smtp.login(smtp_config['smtp_username'], smtp_config['smtp_password'])

    return smtp

def get_smtp_connection(smtp_config, force_new=False):
    """Create or retrieve a thread-local SMTP connection with retry and refresh logic."""
    def connect_with_retry():
        retries = 0
        last_error = None
        while retries < MAX_RETRIES:
            try:
                return create_smtp_connection(smtp_config)
            except (socket.error, smtplib.SMTPException) as e:
                last_error = e
                retries += 1
                backoff = (2 ** retries) + random.random()
                logging.warning(f"SMTP connection attempt {retries} failed: {e}. Retrying in {backoff:.2f}s")
                time.sleep(backoff)
        raise Exception(f"Failed to connect to SMTP after {MAX_RETRIES} retries: {last_error}")
    
    if force_new or not hasattr(thread_local, "smtp") or not hasattr(thread_local, "email_count"):
        if hasattr(thread_local, "smtp"):
            try:
                thread_local.smtp.quit()
            except Exception as e:
                logging.warning(f"Error closing existing SMTP connection: {e}")
        thread_local.email_count = 0
        thread_local.smtp = connect_with_retry()
        return thread_local.smtp

    thread_local.email_count += 1
    if thread_local.email_count >= CONNECTION_REFRESH_COUNT:
        logging.info(f"Refreshing SMTP connection after {CONNECTION_REFRESH_COUNT} emails")
        return get_smtp_connection(smtp_config, force_new=True)

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
            
        # Parse timeout settings with defaults
        config['smtp_timeout'] = int(config.get('smtp_timeout', '60'))  # Default 60 seconds
        
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
        tuple: (email_tasks, original_rows, fieldnames) where:
            - email_tasks is a list of (email, attachments) tuples for sending
            - original_rows is a list of the original CSV rows
            - fieldnames is a list of the CSV column names
    """
    email_tasks = []
    original_rows = []
    
    try:
        # First, filter out comment lines and blank lines
        filtered_lines = []
        with open(mapping_file, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                # Skip blank lines and comment lines
                if not line or line.startswith('#'):
                    continue
                filtered_lines.append(line)
        
        # Use StringIO to create an in-memory file-like object
        from io import StringIO
        csv_data = StringIO('\n'.join(filtered_lines))
        
        # Now read the filtered CSV data
        reader = csv.DictReader(csv_data)
        
        # Clean up fieldnames to remove any BOM characters resulting from Excel conversion to csv
        reader.fieldnames = [field.strip('\ufeff') for field in reader.fieldnames]
        fieldnames = reader.fieldnames
        
        # Verify email column exists
        if email_column not in reader.fieldnames:
            raise ValueError(f"Email column '{email_column}' not found in mapping file. Available columns: {reader.fieldnames}")
        
        logging.info(f"Looking for email column '{email_column}' and attachment columns: {', '.join(attachment_columns)}")
        
        # Verify all attachment columns exist
        missing_columns = [col for col in attachment_columns if col not in reader.fieldnames]
        if missing_columns:
            raise ValueError(f"Attachment column(s) not found in mapping file: {', '.join(missing_columns)}. Available columns: {reader.fieldnames}")
        
        # Read mappings - each row becomes one email task
        row_count = 0
        for row in reader:
            row_count += 1
            email = row[email_column]
            original_rows.append(row)  # Store the original row
            
            if not email or '@' not in email:
                logging.warning(f"Skipping invalid email address in row {row_count}: {email}")
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
        
        logging.info(f"Found CSV columns: {reader.fieldnames}")
        logging.info(f"Found {len(email_tasks)} email tasks to process")
        return email_tasks, original_rows, fieldnames
    
    except Exception as e:
        raise ValueError(f"Error reading mapping file: {str(e)}")

def write_failed_report(failed_tasks, original_fieldnames, output_file=None):
    """Write a CSV report of failed email tasks.
    
    Args:
        failed_tasks (list): List of rows that failed to process
        original_fieldnames (list): Column names from the original CSV
        output_file (str, optional): Output file path. If None, a timestamped filename is generated.
        
    Returns:
        str: Path to the report file, or None if no failed tasks
    """
    # If no failed tasks, don't create a report
    if not failed_tasks:
        logging.info("No failed tasks to report")
        return None
    
    # Create logs directory if it doesn't exist
    os.makedirs('logs', exist_ok=True)
    
    # Generate filename with timestamp if not provided
    if not output_file:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_file = f"logs/failed_emails_{timestamp}.csv"
    
    try:
        # Write the report
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=original_fieldnames)
            
            # Add a comment at the top of the file explaining how to use it
            f.write("# This file contains rows that failed to process.\n")
            f.write("# To retry these emails, use this file as your mapping file:\n")
            f.write(f"# python send_emails_with_pdf_attachments.py <config_file> (with mapping_file = {output_file} in your config)\n\n")
            
            writer.writeheader()
            for row in failed_tasks:
                writer.writerow(row)
                
        logging.info(f"Failed tasks report written to: {output_file}")
        return output_file
        
    except Exception as e:
        logging.error(f"Error writing failed tasks report: {str(e)}")
        return None
    
def validate_file(file_path):
    """
    Validates that a file exists and is a valid PDF document.
    Returns:
        tuple: (is_valid, error_message, file_info)
    """
    try:
        if not os.path.isfile(file_path):
            return False, f"File does not exist or is not a regular file: {file_path}", None
        
        if os.path.getsize(file_path) == 0:
            return False, f"File is empty: {file_path}", None

        # Read file signature for basic PDF check
        with open(file_path, 'rb') as f:
            header = f.read(5)
            if header != b'%PDF-':
                return False, f"File is not a valid PDF (missing %PDF header): {file_path}", None

        # Still collect MIME info for downstream use
        ctype, encoding = mimetypes.guess_type(file_path)
        ctype = ctype or 'application/pdf'  # fallback
        maintype, subtype = ctype.split('/', 1)

        return True, None, {
            'path': file_path,
            'filename': os.path.basename(file_path),
            'ctype': ctype,
            'maintype': maintype,
            'subtype': subtype,
        }

    except Exception as e:
        return False, f"Error checking file: {str(e)}", None

def create_email_message(smtp_config, to_email, subject, body):
    """Create a basic email message with headers and body.
    
    Args:
        smtp_config (dict): SMTP configuration
        to_email (str): Recipient email address
        subject (str): Email subject
        body (str): Email body text
        
    Returns:
        MIMEMultipart: Email message object
    """
    msg = MIMEMultipart()
    # Use smtp_username if auth is enabled, otherwise use from_email from config
    msg['From'] = smtp_config.get('smtp_username') if smtp_config.get('use_auth') else smtp_config['from_email']
    msg['To'] = to_email
    msg['Subject'] = subject
    msg['Date'] = email.utils.formatdate(localtime=True)  # Add proper date header
    msg['Message-ID'] = email.utils.make_msgid(domain=msg['From'].split('@')[1] if '@' in msg['From'] else 'localhost')
    
    # Add body as plain text
    msg.attach(MIMEText(body, 'plain', 'utf-8'))
    
    return msg

def process_binary_attachment(file_path, maintype, subtype):
    """Process a binary file attachment, preserving the original content exactly as-is.
    
    Args:
        file_path (str): Path to the binary file
        maintype (str): MIME maintype
        subtype (str): MIME subtype
        
    Returns:
        tuple: (success, part or None, error message or None)
    """
    try:
        part = MIMEBase(maintype, subtype)
        
        # Read file directly without chunking to preserve exact content
        with open(file_path, 'rb') as f:
            content = f.read()
        
        part.set_payload(content)
        
        # Use standard base64 encoding - preserves original file exactly
        encoders.encode_base64(part)
        
        return True, part, None
    except Exception as e:
        return False, None, str(e)

def process_attachment(attachment_path):
    """Process a single attachment file.
    
    Args:
        attachment_path (str): Path to the attachment file
        
    Returns:
        tuple: (success, part or None, filename or None, error_message or None)
    """
    try:
        # Get file info for attaching
        filename = os.path.basename(attachment_path)
        ctype, encoding = mimetypes.guess_type(attachment_path)
        
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'  # Fallback
        
        maintype, subtype = ctype.split('/', 1)
        
        # Only process PDF files
        if maintype == 'application' and subtype == 'pdf':
            success, part, error = process_binary_attachment(
                attachment_path, maintype, subtype
            )
            if success and part:
                part = add_attachment_headers(part, filename)
                return True, part, filename, None
            else:
                return False, None, filename, f"Error processing PDF file: {error}"
        else:
            # Reject non-PDF files
            return False, None, filename, f"File type not supported: {ctype}. Only PDF files are allowed."
        
    except Exception as e:
        return False, None, os.path.basename(attachment_path), str(e)

def process_all_attachments(msg, attachment_paths, progress_info=None, to_email=None):
    """Process and add all PDF attachments to the email message.
    
    Args:
        msg (MIMEMultipart): Email message object
        attachment_paths (list): List of paths to PDF files
        progress_info (tuple, optional): Tuple of (current, total) for progress reporting
        to_email (str, optional): Recipient email for logging
        
    Returns:
        tuple: (bool, list, list) - success flag, list of attachment names, list of error messages
    """
    attachment_names = []
    attachment_errors = []
    any_attachments_failed = False
    
    # Skip if no attachments
    if not attachment_paths:
        return True, attachment_names, attachment_errors
    
    for attachment_path in attachment_paths:
        # Check if we're being interrupted
        if should_exit.is_set():
            progress_str = f"{progress_info[0]}/{progress_info[1]}" if progress_info else "?"
            logging.info(f"[{progress_str}] Cancelling email to {to_email} due to user interrupt")
            return False, attachment_names, attachment_errors
        
        # Process this PDF attachment
        success, part, filename, error = process_attachment(attachment_path)
        
        if success and part:
            # Add the attachment to the message
            msg.attach(part)
            attachment_names.append(filename)
        else:
            error_msg = f"Error attaching PDF file {filename}: {error}"
            attachment_errors.append(error_msg)
            any_attachments_failed = True
            # Log the error
            logging.error(error_msg)
    
    # If any attachments were supposed to be included but failed, don't send the email
    if attachment_paths and not attachment_names:
        # All attachments failed
        if to_email:
            logging.error(f"Aborting email to {to_email} - all PDF attachments failed to process")
        return False, attachment_names, attachment_errors
    
    # If some but not all attachments failed, and we're strict about it
    if any_attachments_failed:
        if to_email:
            logging.error(f"Aborting email to {to_email} - some PDF attachments failed to process: {'; '.join(attachment_errors)}")
        return False, attachment_names, attachment_errors
    
    return True, attachment_names, attachment_errors

def handle_test_mode(smtp_config, to_email, subject, attachment_names, progress_info=None):
    """Handle test mode by logging instead of sending.
    
    Args:
        smtp_config (dict): SMTP configuration
        to_email (str): Recipient email address
        subject (str): Email subject
        attachment_names (list): List of attachment file names
        progress_info (tuple, optional): Tuple of (current, total) for progress reporting
        
    Returns:
        bool: Success flag
    """
    progress_str = f"[{progress_info[0]}/{progress_info[1]}] " if progress_info else ""
    
    # Check if we should exit early
    if should_exit.is_set():
        logging.info(f"{progress_str}Cancelling email to {to_email} due to user interrupt")
        return False
        
    # Collect all info in a single log message for test mode
    test_info = []
    test_info.append(f"{progress_str}Would send email to {to_email}")
    test_info.append(f"From: {smtp_config.get('smtp_username') if smtp_config.get('use_auth') else smtp_config['from_email']}")
    if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
        test_info.append(f"Bcc: {', '.join(smtp_config['bcc_recipients'])}")
    test_info.append(f"Subject: {subject}")
    
    attachment_str = ", ".join(attachment_names)
    test_info.append(f"Attachments: {attachment_str if attachment_str else 'None'}")
    
    # Log everything in a single line
    logging.info(" | ".join(test_info))
    return True

def add_attachment_headers(part, filename):
    """Add headers to an attachment part.
    
    Args:
        part (MIMEBase): The attachment part
        filename (str): Original filename of the attachment
        
    Returns:
        MIMEBase: The part with headers added
    """
    # Set headers for the attachment
    part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
    
    # Add a Content-ID for potential inline display
    content_id = f"<{filename.replace(' ', '_')}>"
    part.add_header('Content-ID', content_id)
    
    return part

def send_email_with_retry(smtp_config, msg, to_email, progress_info=None):
    """Send an email with retry logic.
    
    Args:
        smtp_config (dict): SMTP configuration
        msg (MIMEMultipart): Email message object
        to_email (str): Recipient email address
        progress_info (tuple, optional): Tuple of (current, total) for progress reporting
        
    Returns:
        tuple: (bool, float) - success flag, elapsed time
    """
    start_time = time.time()
    
    # Format progress info for logging
    progress_str = f"[{progress_info[0]}/{progress_info[1]}] " if progress_info else ""
    
    # Send email with retry logic
    retries = 0
    while retries < MAX_RETRIES:
        # Check if we should exit early
        if should_exit.is_set():
            logging.info(f"{progress_str}Cancelling email to {to_email} due to user interrupt")
            return False, 0
            
        try:
            # Get a connection (might be new or existing)
            smtp = get_smtp_connection(smtp_config)
            
            # Use the more reliable sendmail method instead of send_message
            # This gives us more control over the SMTP transaction
            from_addr = smtp_config.get('smtp_username') if smtp_config.get('use_auth') else smtp_config['from_email']
            to_addrs = [to_email]
            
            # Add BCCs to recipient list if configured
            if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
                to_addrs.extend(smtp_config['bcc_recipients'])
            
            # Convert message to string with proper encoding
            msg_str = msg.as_string()
            
            # Send the email using raw SMTP commands for more reliable delivery
            smtp.sendmail(from_addr, to_addrs, msg_str)
            
            elapsed_time = time.time() - start_time
            return True, elapsed_time
        
        except (smtplib.SMTPException, socket.error, ConnectionError, OSError) as e:
            retries += 1
            elapsed_time = time.time() - start_time
            
            # Check if we should exit early
            if should_exit.is_set():
                logging.info(f"{progress_str}Cancelling email to {to_email} due to user interrupt")
                return False, elapsed_time
                
            # Determine if we should retry
            if retries < MAX_RETRIES:
                backoff_time = (2 ** retries) + random.random()
                logging.warning(f"{progress_str}Error sending to {to_email} (attempt {retries}/{MAX_RETRIES}) | Error: {str(e)} | Retrying in {backoff_time:.2f}s")
                time.sleep(backoff_time)
                
                # Force a new connection on retry
                try:
                    smtp = get_smtp_connection(smtp_config, force_new=True)
                except Exception as conn_err:
                    logging.error(f"{progress_str}Failed to refresh connection for {to_email} | Error: {str(conn_err)}")
            else:
                logging.error(f"{progress_str}Failed to send to {to_email} | Time: {elapsed_time:.2f}s | Error: {str(e)} | Max retries exceeded")
                return False, elapsed_time

def send_email(smtp_config, to_email, subject, body, attachment_paths, test_mode=False, progress=None):
    """Send an email with file attachments with retry logic.
    
    This function coordinates the entire email sending process by:
    1. Creating the email message
    2. Processing attachments
    3. Either logging (test mode) or sending the email
    
    Args:
        smtp_config (dict): SMTP configuration dictionary
        to_email (str): Recipient email address
        subject (str): Email subject
        body (str): Email body text
        attachment_paths (list): List of paths to attachment files
        test_mode (bool): If True, don't actually send emails
        progress (tuple): Tuple of (current, total) for progress reporting
        
    Returns:
        bool: Whether the email was sent successfully
    """
    # Check if we should exit early
    if should_exit.is_set():
        logging.info(f"{progress[0] if progress else '?'}/{progress[1] if progress else '?'} Cancelling email to {to_email} due to user interrupt")
        return False
    
    # 1. Create basic message with headers and body
    msg = create_email_message(
        smtp_config, 
        to_email, 
        subject, 
        body
    )
    
    # 2. Process all attachments
    attachment_success, attachment_names, attachment_errors = process_all_attachments(
        msg, 
        attachment_paths,
        progress_info=progress,
        to_email=to_email
    )
    
    # Exit if attachment processing failed
    if not attachment_success:
        return False
    
    # 3. Handle test mode if enabled
    if test_mode:
        return handle_test_mode(
            smtp_config,
            to_email,
            subject,
            attachment_names,
            progress_info=progress
        )
    
    # 4. Send the email
    success, elapsed_time = send_email_with_retry(
        smtp_config,
        msg,
        to_email,
        progress_info=progress
    )
    
    # 5. Log success
    if success:
        progress_info = f"[{progress[0]}/{progress[1]}] " if progress else ""
        log_parts = []
        log_parts.append(f"{progress_info}Email sent to {to_email}")
        log_parts.append(f"Attachments: {len(attachment_names)}")
        log_parts.append(f"Time: {elapsed_time:.2f}s")
        
        # Log everything in a single line
        logging.info(" | ".join(log_parts))
    
    return success

def process_email(args):
    """Process a single email send operation."""
    
    # Check if we should exit early
    if should_exit.is_set():
        return False, 0, args[6]  # Return failure, zero time, and row index
        
    try:
        smtp_config, recipient_email, attachment_paths, config, current_count, total_emails, row_index = args
        email_start_time = time.time()
        success = send_email(
            smtp_config=smtp_config,
            to_email=recipient_email,
            subject=config['email_subject'],
            body=config['email_body'],
            attachment_paths=attachment_paths,
            test_mode=config['test_mode'],
            progress=(current_count, total_emails)
        )
        return success, time.time() - email_start_time, row_index
    except Exception as e:
        # Log full stack trace for unexpected exceptions
        stack_trace = traceback.format_exc()
        logging.error(f"Exception in process_email: {str(e)}\n{stack_trace}")
        return False, 0, args[6]  # Return failure with the row index

def setup_environment(config_path):
    """Set up the environment by reading config and initializing logging.
    
    Args:
        config_path (str): Path to configuration file
        
    Returns:
        tuple: (config dict, log file path)
    """
    # Read configuration first (before setting up logging)
    config = read_config(config_path)
    
    # Set up logging
    log_file = setup_logging()
    logging.info(f"Starting email sending process with config file: {config_path}")
    
    # Verify input directory exists
    input_dir = config['input_directory']
    if not os.path.isdir(input_dir):
        raise ValueError(f"Input directory does not exist: {input_dir}")
        
    return config, log_file

def read_and_validate_tasks(config):
    """Read the mapping file and validate the email tasks.
    
    Args:
        config (dict): Configuration dictionary
        
    Returns:
        tuple: (email_tasks, original_rows, fieldnames, total_emails, total_attachments)
    """
    logging.info("\nReading mapping file...")
    email_tasks, original_rows, fieldnames = read_mapping_file(
        config['mapping_file'],
        config['email_column'],
        config['attachment_columns']
    )
    
    if not email_tasks:
        raise ValueError("No valid mappings found in mapping file")
    
    # Calculate statistics
    total_emails = len(email_tasks)
    total_attachments = sum(len(files) for _, files in email_tasks)
    
    # Log basic info
    logging.info(f"\nProcessing {total_emails} emails with {total_attachments} total attachments")
    logging.info(f"Using {int(config.get('max_threads', '4'))} threads for parallel processing")
    logging.info(f"SMTP Server: {config['smtp_server']}:{config['smtp_port']}")
    logging.info(f"From: {config.get('smtp_username') if config.get('use_auth') else config['from_email']}")
    logging.info(f"TLS: {'Enabled' if config.get('use_tls') else 'Disabled'}")
    logging.info(f"Authentication: {'Enabled' if config.get('use_auth') else 'Disabled'}")
    logging.info(f"Attachment columns: {', '.join(config['attachment_columns'])}")
    
    if config['test_mode']:
        logging.info("(TEST MODE - Emails will not be sent)")
    if config['bcc_recipients']:
        logging.info(f"BCC recipients: {', '.join(config['bcc_recipients'])}")
    
    return email_tasks, original_rows, fieldnames, total_emails, total_attachments

def verify_attachment_files(email_tasks, input_dir, total_emails):
    """Verify that all attachment files exist and are valid PDF files.
    
    Args:
        email_tasks (list): List of (email, attachments) tuples
        input_dir (str): Directory containing attachment files
        total_emails (int): Total number of emails to process
        
    Returns:
        dict: Dictionary mapping filenames to whether they exist and are valid PDFs
        
    Raises:
        ValueError: If any attachment files don't exist or are not valid PDFs
    """
    logging.info("\nVerifying all PDF attachment files exist and are valid...")
    
    found_files = {}  # Track files that exist so we don't check multiple times
    invalid_attachments = []  # Track all invalid attachment files
    not_found_count = 0
    invalid_type_count = 0
    
    # Check each attachment file
    for index, (recipient_email, attachment_files) in enumerate(email_tasks):
        # Check for exit signal
        if should_exit.is_set():
            logging.info("Verification interrupted by user. Exiting...")
            return None
            
        current_count = index + 1
        
        if not attachment_files:
            continue  # Skip checking if no attachments
            
        # Check which attachment files exist
        row_invalid_files = []
        
        for attachment_file in attachment_files:
            file_path = os.path.join(input_dir, attachment_file)
            
            # Check if we already know if this file exists
            if attachment_file in found_files:
                if not found_files[attachment_file]:  # File is known to not exist
                    row_invalid_files.append(f"{attachment_file} (file not found)")
            else:
                # Check if the file exists and is a PDF
                is_valid, error_message, _ = validate_file(file_path)
                
                if is_valid:
                    found_files[attachment_file] = True
                else:
                    found_files[attachment_file] = False
                    if "not a PDF document" in error_message:
                        invalid_type_count += 1
                    else:
                        not_found_count += 1
                    row_invalid_files.append(f"{attachment_file} ({error_message})")
        
        # If any attachment files were invalid, log it and add to invalid_attachments list
        if row_invalid_files:
            logging.error(f"[{current_count}/{total_emails}] Email to {recipient_email} | Row {index+1} | Invalid files: {', '.join(row_invalid_files)}")
            invalid_attachments.append(f"Row {index+1} - Email to {recipient_email} has invalid file(s): {', '.join(row_invalid_files)}")
    
    # If any attachments are invalid, abort the process
    if invalid_attachments:
        for msg in invalid_attachments:
            logging.error(msg)
        
        error_msg = []
        if not_found_count > 0:
            error_msg.append(f"{not_found_count} missing file(s)")
        if invalid_type_count > 0:
            error_msg.append(f"{invalid_type_count} non-PDF file(s)")
            
        raise ValueError(f"Aborting: Found {', '.join(error_msg)}. Please fix attachment issues before running again.")
    
    return found_files

def prepare_email_tasks(email_tasks, config):
    """Prepare the email tasks for processing.
    
    Args:
        email_tasks (list): List of (email, attachments) tuples
        config (dict): Configuration dictionary
        
    Returns:
        list: List of task tuples ready for processing
    """
    processing_tasks = []
    current_count = 0
    input_dir = config['input_directory']
    
    # Process each email task
    for index, (recipient_email, attachment_files) in enumerate(email_tasks):
        # Check for interruption
        if should_exit.is_set():
            break
            
        current_count += 1
        
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
        
        # If no attachment files specified, send email without attachments
        if not attachment_files:
            # Add task with empty attachments list
            processing_tasks.append((smtp_config, recipient_email, [], config, current_count, len(email_tasks), index))
            continue
        
        # Collect valid file paths for attachments
        valid_file_paths = []
        for attachment_file in attachment_files:
            file_path = os.path.join(input_dir, attachment_file)
            valid_file_paths.append(file_path)
        
        # Add task for this email with its attachments
        processing_tasks.append((smtp_config, recipient_email, valid_file_paths, config, current_count, len(email_tasks), index))
    
    return processing_tasks

def process_emails_in_parallel(processing_tasks, max_threads):
    """Process emails in parallel using a thread pool.
    
    Args:
        processing_tasks (list): List of task tuples
        max_threads (int): Maximum number of threads to use
        
    Returns:
        tuple: (success_count, skipped_count, failed_rows, total_email_time)
    """
    success_count = 0
    skipped_count = 0
    total_email_time = 0
    failed_rows = []
    
    # Process emails in parallel using thread pool
    with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
        # Use a dict to track which future corresponds to which row
        future_to_index = {}
        futures = []
        
        # Submit all tasks
        for task in processing_tasks:
            # Check if we've been interrupted
            if should_exit.is_set():
                break
                
            future = executor.submit(process_email, task)
            future_to_index[future] = task[6]  # Store the row index
            futures.append(future)
        
        # Process results as they complete
        try:
            for future in concurrent.futures.as_completed(futures):
                # This will raise any exceptions from the task
                if should_exit.is_set():
                    # Don't wait for all tasks, just process the ones that have completed
                    if future.done():
                        row_index = future_to_index[future]
                        try:
                            success, email_time, _ = future.result()
                            if success:
                                success_count += 1
                                total_email_time += email_time
                            else:
                                skipped_count += 1
                                failed_rows.append(row_index)
                        except Exception as e:
                            stack_trace = traceback.format_exc()
                            logging.error(f"Error processing email: {str(e)}\n{stack_trace}")
                            skipped_count += 1
                            failed_rows.append(row_index)
                else:
                    row_index = future_to_index[future]
                    try:
                        success, email_time, _ = future.result()
                        if success:
                            success_count += 1
                            total_email_time += email_time
                        else:
                            skipped_count += 1
                            failed_rows.append(row_index)
                    except Exception as e:
                        stack_trace = traceback.format_exc()
                        logging.error(f"Error processing email: {str(e)}\n{stack_trace}")
                        skipped_count += 1
                        failed_rows.append(row_index)
        except KeyboardInterrupt:
            # This is a backup in case the signal handler doesn't catch it
            logging.info("\nKeyboard interrupt detected during email processing.")
            should_exit.set()
        
        # If interrupted, cancel any pending futures
        if should_exit.is_set():
            logging.info("Cancelling any pending email tasks...")
            cancelled_count = 0
            for future in futures:
                if not future.done():
                    future.cancel()
                    cancelled_count += 1
            
            if cancelled_count > 0:
                logging.info(f"Cancelled {cancelled_count} pending email tasks")
    
    return success_count, skipped_count, failed_rows, total_email_time

def generate_report(success_count, skipped_count, total_emails, total_attachments, 
                   total_email_time, failed_rows, fieldnames, total_time, log_file):
    """Generate a final report of the email sending process.
    
    Args:
        success_count (int): Number of successfully sent emails
        skipped_count (int): Number of skipped/failed emails
        total_emails (int): Total number of emails to process
        total_attachments (int): Total number of attachments
        total_email_time (float): Total time spent sending emails
        failed_rows (list): List of failed row indices
        fieldnames (list): CSV field names
        total_time (float): Total processing time
        log_file (str): Path to log file
        
    Returns:
        str or None: Path to failed report file, if any
    """
    # Check if we were interrupted
    if should_exit.is_set():
        logging.info("\nEmail sending process was interrupted by user.")
    
    # Write report of failed rows
    failed_report = None
    if failed_rows:
        failed_report = write_failed_report(failed_rows, fieldnames)
    
    # Calculate statistics
    avg_email_time = total_email_time / success_count if success_count > 0 else 0
    
    # Calculate correct number of unprocessed emails
    processed_emails = success_count + skipped_count
    unprocessed_emails = total_emails - processed_emails
    
    # Print summary
    logging.info("\nSummary:")
    logging.info(f"Total time: {total_time:.2f} seconds")
    logging.info(f"Average time per email: {avg_email_time:.2f} seconds")
    logging.info(f"Total rows: {total_emails}")
    logging.info(f"Total attachments: {total_attachments}")
    logging.info(f"Successfully sent: {success_count}")
    logging.info(f"Emails skipped/failed: {skipped_count}")
    
    if should_exit.is_set():
        logging.info(f"Process interrupted: {unprocessed_emails} emails not processed")
    
    logging.info(f"Log file: {log_file}")
    
    if failed_report:
        logging.info(f"Failed tasks report: {failed_report}")
        
    return failed_report

def cleanup_resources():
    """Clean up any remaining resources before exiting."""
    # Clean up any remaining SMTP connections
    if hasattr(thread_local, "smtp"):
        try:
            thread_local.smtp.quit()
        except Exception as e:
            logging.warning(f"Error closing SMTP connection: {str(e)}")

def main():
    """Main function to coordinate the email sending process."""
    
    if len(sys.argv) != 2:
        print("Usage: python send_emails_with_pdf_attachments.py <config_file>")
        sys.exit(1)
    
    try:
        # 1. Setup environment (config and logging)
        start_time = time.time()
        config, log_file = setup_environment(sys.argv[1])
        
        # 2. Read and validate email tasks
        email_tasks, original_rows, fieldnames, total_emails, total_attachments = read_and_validate_tasks(config)
        
        # 3. Verify all attachment files exist
        verify_attachment_files(email_tasks, config['input_directory'], total_emails)
        
        # Exit if interrupted
        if should_exit.is_set():
            logging.info("Process interrupted by user before sending emails. Exiting...")
            return
            
        # 4. Prepare email tasks for processing
        processing_tasks = prepare_email_tasks(email_tasks, config)
        
        # 5. Process emails in parallel using a thread pool
        max_threads = int(config.get('max_threads', '4'))
        success_count, skipped_count, failed_row_indices, total_email_time = process_emails_in_parallel(
            processing_tasks, max_threads
        )
        
        # Get the actual failed rows from the original rows
        failed_rows = [original_rows[i] for i in failed_row_indices]
        
        # 6. Generate final report
        total_time = time.time() - start_time
        generate_report(
            success_count, skipped_count, total_emails, total_attachments,
            total_email_time, failed_rows, fieldnames, total_time, log_file
        )
        
    except KeyboardInterrupt:
        # This is a backup in case the signal handler doesn't catch it
        logging.error("\nScript interrupted with keyboard interrupt (Ctrl+C).")
        should_exit.set()
    except Exception as e:
        stack_trace = traceback.format_exc()
        logging.error(f"\nUnexpected error: {str(e)}\n{stack_trace}")
        sys.exit(1)
    finally:
        # 7. Clean up resources
        cleanup_resources()
        
        # Normal exit if interrupted, error code otherwise
        if should_exit.is_set():
            logging.info("Script terminated due to user interrupt.")
            sys.exit(0)
        elif 'success_count' in locals() and 'skipped_count' in locals() and 'total_emails' in locals():
            if success_count + skipped_count < total_emails:
                logging.error("Script terminated abnormally, not all emails were processed.")
                sys.exit(1)

if __name__ == '__main__':
    main() 