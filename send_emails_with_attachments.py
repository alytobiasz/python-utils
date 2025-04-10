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

# Standard library imports
import concurrent.futures
import csv
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
from datetime import datetime
from email import encoders
from email.mime.application import MIMEApplication
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path


# Thread-local storage for SMTP connections
thread_local = threading.local()

# Maximum number of connection retries
MAX_RETRIES = 3

# Connection refresh settings - refresh connection every X emails
CONNECTION_REFRESH_COUNT = 20  # Refresh SMTP connection after this many emails

# Global flag to track if script should exit
should_exit = False

# Set up signal handler for graceful shutdown
def signal_handler(sig, frame):
    """Handle keyboard interrupt and other signals to gracefully shut down."""
    global should_exit
    logging.info("\nKeyboard interrupt detected. Shutting down gracefully...")
    logging.info("Please wait for current tasks to finish (this may take a moment)...")
    
    # Set flag for threads to check
    should_exit = True
    
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

def get_smtp_connection(smtp_config, force_new=False):
    """Get or create an SMTP connection for the current thread.
    
    Args:
        smtp_config (dict): SMTP configuration dictionary
        force_new (bool): If True, create a new connection even if one exists
        
    Returns:
        smtplib.SMTP: The SMTP connection
    
    Raises:
        Exception: If connection fails after MAX_RETRIES attempts
    """
    # If force_new or no connection exists, create a new one
    if force_new or not hasattr(thread_local, "smtp") or not hasattr(thread_local, "email_count"):
        # Close existing connection if it exists
        if hasattr(thread_local, "smtp"):
            try:
                thread_local.smtp.quit()
            except:
                pass  # Ignore errors when closing
        
        # Initialize or reset email counter for this thread
        thread_local.email_count = 0
        
        # Retry logic for creating the connection
        retries = 0
        last_error = None
        
        while retries < MAX_RETRIES:
            try:
                # Create new SMTP connection for this thread
                smtp = smtplib.SMTP(smtp_config['smtp_server'], smtp_config['smtp_port'], timeout=30)
                
                # Add some basic error handling
                smtp.set_debuglevel(0)
                
                # TLS if needed
                if smtp_config.get('use_tls', False):
                    smtp.starttls()
                
                # Authentication if needed
                if smtp_config.get('use_auth', False):
                    smtp.login(smtp_config['smtp_username'], smtp_config['smtp_password'])
                
                thread_local.smtp = smtp
                return smtp
            
            except (socket.error, smtplib.SMTPException) as e:
                last_error = e
                retries += 1
                
                # Add randomized exponential backoff
                backoff_time = (2 ** retries) + random.random()
                logging.warning(f"SMTP connection attempt {retries} failed: {str(e)}. Retrying in {backoff_time:.2f} seconds...")
                time.sleep(backoff_time)
        
        # If we get here, we've exhausted our retries
        raise Exception(f"Failed to establish SMTP connection after {MAX_RETRIES} attempts: {str(last_error)}")
    
    # Check if we need to refresh the connection
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
            f.write(f"# python send_emails_with_attachments.py <config_file> (with mapping_file = {output_file} in your config)\n\n")
            
            writer.writeheader()
            for row in failed_tasks:
                writer.writerow(row)
                
        logging.info(f"Failed tasks report written to: {output_file}")
        return output_file
        
    except Exception as e:
        logging.error(f"Error writing failed tasks report: {str(e)}")
        return None

def send_email(smtp_config, to_email, subject, body, attachment_paths, test_mode=False, progress=None):
    """Send an email with file attachments with retry logic."""
    global should_exit
    start_time = time.time()
    
    # Check if we should exit early
    if should_exit:
        logging.info(f"{progress[0] if progress else '?'}/{progress[1] if progress else '?'} Cancelling email to {to_email} due to user interrupt")
        return False
    
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
    attachment_errors = []
    for attachment_path in attachment_paths:
        if should_exit:
            logging.info(f"{progress[0] if progress else '?'}/{progress[1] if progress else '?'} Cancelling email to {to_email} due to user interrupt")
            return False
    
        try:
            filename = os.path.basename(attachment_path)
            ctype, encoding = mimetypes.guess_type(attachment_path)
    
            if ctype is None or encoding is not None:
                ctype = 'application/octet-stream'  # Fallback
    
            maintype, subtype = ctype.split('/', 1)
    
            with open(attachment_path, 'rb') as f:
                if maintype == 'text':
                    part = MIMEText(f.read().decode('utf-8'), _subtype=subtype)
                elif maintype == 'image':
                    part = MIMEImage(f.read(), _subtype=subtype)
                elif maintype == 'audio':
                    part = MIMEAudio(f.read(), _subtype=subtype)
                else:
                    part = MIMEBase(maintype, subtype)
                    part.set_payload(f.read())
                    encoders.encode_base64(part)
    
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)
            attachment_names.append(filename)
    
        except Exception as e:
            error_msg = f"Error attaching file {filename}: {e}"
            attachment_errors.append(error_msg)
            # Continue with other attachments
    
    progress_info = f"[{progress[0]}/{progress[1]}] " if progress else ""
    attachment_str = ", ".join(attachment_names)
    
    if test_mode:
        # Check if we should exit early
        if should_exit:
            logging.info(f"{progress_info}Cancelling email to {to_email} due to user interrupt")
            return False
            
        # Collect all info in a single log message for test mode
        test_info = []
        test_info.append(f"{progress_info}Would send email to {to_email}")
        test_info.append(f"From: {msg['From']}")
        if 'bcc_recipients' in smtp_config and smtp_config['bcc_recipients']:
            test_info.append(f"Bcc: {', '.join(smtp_config['bcc_recipients'])}")
        test_info.append(f"Subject: {subject}")
        test_info.append(f"Attachments: {attachment_str if attachment_str else 'None'}")
        
        # If there were attachment errors, add those
        if attachment_errors:
            test_info.append(f"Attachment Errors: {'; '.join(attachment_errors)}")
        
        # Log everything in a single line
        logging.info(" | ".join(test_info))
        return True
    
    # Send email with retry logic
    retries = 0
    while retries < MAX_RETRIES:
        # Check if we should exit early
        if should_exit:
            logging.info(f"{progress_info}Cancelling email to {to_email} due to user interrupt")
            return False
            
        try:
            if not test_mode:
                # Get a connection (might be new or existing)
                smtp = get_smtp_connection(smtp_config)
                smtp.send_message(msg)
            
            elapsed_time = time.time() - start_time
            
            # Create a comprehensive success log message
            log_parts = []
            log_parts.append(f"{progress_info}Email sent to {to_email}")
            log_parts.append(f"Attachments: {len(attachment_names)}")
            log_parts.append(f"Time: {elapsed_time:.2f}s")
            
            # If there were attachment errors, add those
            if attachment_errors:
                log_parts.append(f"Attachment Errors: {'; '.join(attachment_errors)}")
            
            # Log everything in a single line
            logging.info(" | ".join(log_parts))
            return True
        
        except (smtplib.SMTPException, socket.error, ConnectionError, OSError) as e:
            retries += 1
            elapsed_time = time.time() - start_time
            
            # Check if we should exit early
            if should_exit:
                logging.info(f"{progress_info}Cancelling email to {to_email} due to user interrupt")
                return False
                
            # Determine if we should retry
            if retries < MAX_RETRIES:
                backoff_time = (2 ** retries) + random.random()
                logging.warning(f"{progress_info}Error sending to {to_email} (attempt {retries}/{MAX_RETRIES}) | Error: {str(e)} | Retrying in {backoff_time:.2f}s")
                time.sleep(backoff_time)
                
                # Force a new connection on retry
                try:
                    smtp = get_smtp_connection(smtp_config, force_new=True)
                except Exception as conn_err:
                    logging.error(f"{progress_info}Failed to refresh connection for {to_email} | Error: {str(conn_err)}")
            else:
                logging.error(f"{progress_info}Failed to send to {to_email} | Time: {elapsed_time:.2f}s | Error: {str(e)} | Max retries exceeded")
                return False

def process_email(args):
    """Process a single email send operation."""
    global should_exit
    
    # Check if we should exit early
    if should_exit:
        return False, 0, args[6]  # Return failure, zero time, and row index
        
    smtp_config, email, attachment_paths, config, current_count, total_emails, row_index = args
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
    return success, time.time() - email_start_time, row_index

def main():
    global should_exit
    
    if len(sys.argv) != 2:
        print("Usage: python send_emails_with_attachments.py <config_file>")
        sys.exit(1)
    
    # Set up a main thread exception handler to catch keyboard interrupts
    executor = None
    
    try:
        # Read configuration first (before setting up logging)
        config = read_config(sys.argv[1])
        
        # Set up logging
        log_file = setup_logging()
        logging.info(f"Starting email sending process with config file: {sys.argv[1]}")
        
        start_time = time.time()
        total_email_time = 0
        input_dir = config['input_directory']
        
        # Get number of threads from config or use default
        max_threads = int(config.get('max_threads', '4'))
        
        # Verify input directory exists
        if not os.path.isdir(input_dir):
            raise ValueError(f"Input directory does not exist: {input_dir}")
        
        # Read mapping file
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
        
        # Process attachment files
        success_count = 0
        skipped_count = 0
        not_found_count = 0
        found_files = {}  # Track files that exist so we don't check multiple times
        failed_rows = []  # Store rows that failed to process
        skipped_indices = set()  # Keep track of skipped rows due to missing attachments
        missing_attachments = []  # Track all missing attachment files
        
        logging.info(f"\nProcessing {total_emails} emails with {total_attachments} total attachments")
        logging.info(f"Using {max_threads} threads for parallel processing")
        logging.info(f"SMTP Server: {config['smtp_server']}:{config['smtp_port']}")
        logging.info(f"From: {config.get('smtp_username') if config.get('use_auth') else config['from_email']}")
        logging.info(f"TLS: {'Enabled' if config.get('use_tls') else 'Disabled'}")
        logging.info(f"Authentication: {'Enabled' if config.get('use_auth') else 'Disabled'}")
        logging.info(f"Attachment columns: {', '.join(config['attachment_columns'])}")
        if config['test_mode']:
            logging.info("(TEST MODE - Emails will not be sent)")
        if config['bcc_recipients']:
            logging.info(f"BCC recipients: {', '.join(config['bcc_recipients'])}")
        
        logging.info("\nVerifying all attachment files exist...")
        # First verify all attachment files exist
        for index, (email, attachment_files) in enumerate(email_tasks):
            # Check for exit signal
            if should_exit:
                logging.info("Verification interrupted by user. Exiting...")
                return
                
            current_count = index + 1
            
            if not attachment_files:
                continue  # Skip checking if no attachments
                
            # Check which attachment files exist
            invalid_files = []
            
            for attachment_file in attachment_files:
                file_path = os.path.join(input_dir, attachment_file)
                
                # Check if we already know if this file exists
                if attachment_file in found_files:
                    if not found_files[attachment_file]:  # File is known to not exist
                        invalid_files.append(attachment_file)
                else:
                    # Check if file exists and store result for future use
                    if os.path.isfile(file_path):
                        found_files[attachment_file] = True
                    else:
                        invalid_files.append(attachment_file)
                        found_files[attachment_file] = False
                        not_found_count += 1
            
            # If any attachment files were not found, log it and add to missing_attachments list
            if invalid_files:
                logging.error(f"[{current_count}/{total_emails}] Email to {email} | Row {index+1} | Missing files: {', '.join(invalid_files)}")
                missing_attachments.append(f"Row {index+1} - Email to {email} has missing attachment file(s): {', '.join(invalid_files)}")
        
        # If any attachments are missing, abort the process
        if missing_attachments:
            raise ValueError(f"Aborting: Found {not_found_count} missing attachment file(s). Please fix missing attachments before running again.")
        
        # Check if we should exit
        if should_exit:
            logging.info("Process interrupted by user before sending emails. Exiting...")
            return
            
        # Prepare email tasks for processing
        processing_tasks = []
        current_count = 0
        
        # Process each email task
        for index, (email, attachment_files) in enumerate(email_tasks):
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
                processing_tasks.append((smtp_config, email, [], config, current_count, total_emails, index))
                continue
            
            # At this point, we know all attachments exist, so just collect valid file paths
            valid_file_paths = []
            for attachment_file in attachment_files:
                file_path = os.path.join(input_dir, attachment_file)
                valid_file_paths.append(file_path)
            
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
            processing_tasks.append((smtp_config, email, valid_file_paths, config, current_count, total_emails, index))
        
        # Process emails in parallel using thread pool
        with concurrent.futures.ThreadPoolExecutor(max_workers=max_threads) as executor:
            # Store executor in global variable so it can be shut down by signal handler
            active_executor = executor
            
            # Use a dict to track which future corresponds to which row
            future_to_index = {}
            futures = []
            
            # Submit all tasks
            for task in processing_tasks:
                # Check if we've been interrupted
                if should_exit:
                    break
                    
                future = executor.submit(process_email, task)
                future_to_index[future] = task[6]  # Store the row index
                futures.append(future)
            
            # Process results as they complete
            try:
                for future in concurrent.futures.as_completed(futures):
                    # This will raise any exceptions from the task
                    if should_exit:
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
                                    failed_rows.append(original_rows[row_index])
                            except Exception as e:
                                logging.error(f"Error processing email: {e}")
                                skipped_count += 1
                                failed_rows.append(original_rows[row_index])
                    else:
                        row_index = future_to_index[future]
                        try:
                            success, email_time, _ = future.result()
                            if success:
                                success_count += 1
                                total_email_time += email_time
                            else:
                                skipped_count += 1
                                failed_rows.append(original_rows[row_index])
                        except Exception as e:
                            logging.error(f"Error processing email: {e}")
                            skipped_count += 1
                            failed_rows.append(original_rows[row_index])
            except KeyboardInterrupt:
                # This is a backup in case the signal handler doesn't catch it
                logging.info("\nKeyboard interrupt detected during email processing.")
                should_exit = True
            
            # If interrupted, cancel any pending futures
            if should_exit:
                logging.info("Cancelling any pending email tasks...")
                cancelled_count = 0
                for future in futures:
                    if not future.done():
                        future.cancel()
                        cancelled_count += 1
                
                if cancelled_count > 0:
                    logging.info(f"Cancelled {cancelled_count} pending email tasks")
        
        # Check if we were interrupted
        if should_exit:
            logging.info("\nEmail sending process was interrupted by user.")
        
        # Write report of failed rows
        if failed_rows:
            failed_report = write_failed_report(failed_rows, fieldnames)
        else:
            failed_report = None
        
        # Print summary
        total_time = time.time() - start_time
        avg_email_time = total_email_time / success_count if success_count > 0 else 0
        
        # Calculate correct number of unprocessed emails
        processed_emails = success_count + skipped_count
        unprocessed_emails = total_emails - processed_emails
        
        logging.info("\nSummary:")
        logging.info(f"Total time: {total_time:.2f} seconds")
        logging.info(f"Average time per email: {avg_email_time:.2f} seconds")
        logging.info(f"Total rows: {total_emails}")
        logging.info(f"Total attachments: {total_attachments}")
        logging.info(f"Successfully sent: {success_count}")
        logging.info(f"Files not found: {not_found_count}")
        logging.info(f"Emails skipped/failed: {skipped_count}")
        if should_exit:
            logging.info(f"Process interrupted: {unprocessed_emails} emails not processed")
        logging.info(f"Log file: {log_file}")
        if failed_report:
            logging.info(f"Failed tasks report: {failed_report}")
        
    except KeyboardInterrupt:
        # This is a backup in case the signal handler doesn't catch it
        logging.error("\nScript interrupted with keyboard interrupt (Ctrl+C).")
        should_exit = True
    except Exception as e:
        logging.error(f"\nError: {str(e)}")
        sys.exit(1)
    finally:
        # Clean up any remaining SMTP connections
        if hasattr(thread_local, "smtp"):
            try:
                thread_local.smtp.quit()
            except:
                pass
        
        # Normal exit if interrupted, error code otherwise
        if should_exit:
            logging.info("Script terminated due to user interrupt.")
            sys.exit(0)
        elif 'success_count' in locals() and 'skipped_count' in locals() and 'total_emails' in locals():
            if success_count + skipped_count < total_emails:
                logging.error("Script terminated abnormally, not all emails were processed.")
                sys.exit(1)

if __name__ == '__main__':
    main() 