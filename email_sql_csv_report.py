"""
Database Query to CSV and Email

This script:
1. Connects to a database
2. Runs a SQL query
3. Exports the results to a CSV file
4. Emails the CSV file to specified recipients

Requirements:
    - Python 3.6+
    - Required packages:
        - sqlite3 (built-in) or another DB driver depending on your database
        - pandas (for easy data handling and CSV export)
        - smtplib (built-in for sending emails)
        - email (built-in for crafting email messages)
        - For MS SQL Server: pyodbc
        - For Oracle: oracledb

Usage:
    python email_sql_csv_report.py <config_file>

Example config file:

# Database settings
db_type = sqlite  # sqlite, oracle, mssql

# For SQLite:
# db_connection = mydatabase.db  # Path to the database file

# For Oracle:
# db_connection = localhost:1521/ORCLPDB  # Format: host:port/service_name
# db_user = oracle_user
# db_password = oracle_password

# For MS SQL Server:
# db_server = server_name  # Hostname or IP address
# db_port = 1433  # Default SQL Server port
# db_name = database_name  # Name of the database
# db_driver = ODBC Driver 17 for SQL Server  # ODBC driver name installed on your system
# use_windows_auth = true  # Set to true for Windows Authentication (if false, provide db_user and db_password)

# Alternative for MS SQL Server: Provide a complete connection string (this overrides individual settings above):
# db_connection = Driver={ODBC Driver 17 for SQL Server};Server=server_name;Database=database_name;Trusted_Connection=yes;

# Query settings
query_file = query.sql  # File containing the SQL query
csv_output_dir = path/to/directory  # Where to save the CSV file

# Email settings
smtp_server = smtp.gmail.com
smtp_port = 587
use_tls = true
use_auth = true
smtp_username = your.email@gmail.com
smtp_password = your_app_password
from_email = sender@example.com
recipients = recipient1@example.com, recipient2@example.com
email_subject = Database Query Results
email_body = Please find attached the results of the database query.
"""

import sys
import os
import smtplib
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import logging

# Import utility functions
from utils import connect_to_database, run_query, export_to_csv, read_config

# Ensure logs directory exists
os.makedirs('logs', exist_ok=True)

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(f"logs/email_sql_csv_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    ]
)
logger = logging.getLogger(__name__)

def send_email(config, attachment_file):
    """
    Send an email with the CSV file attached.
    
    Args:
        config: Configuration dictionary
        attachment_file: Path to the CSV file to attach
    """
    # Email server settings
    smtp_server = config.get('smtp_server', '')
    smtp_port = int(config.get('smtp_port', 25))
    use_tls = config.get('use_tls', '').lower() == 'true'
    use_auth = config.get('use_auth', '').lower() == 'true'
    smtp_username = config.get('smtp_username', '')
    smtp_password = config.get('smtp_password', '')
    
    # Email content
    from_email = config.get('from_email', '')
    recipients_str = config.get('recipients', '')
    recipients = [r.strip() for r in recipients_str.split(',') if r.strip()]
    subject = config.get('email_subject', 'Database Query Results')
    body = config.get('email_body', 'Please find attached the results of the database query.')
    
    if not recipients:
        logger.error("No recipients specified.")
        raise ValueError("No recipients specified in the configuration.")
    
    try:
        # Create message
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = ', '.join(recipients)
        msg['Subject'] = subject
        
        # Attach body text
        msg.attach(MIMEText(body, 'plain'))
        
        # Attach CSV file
        with open(attachment_file, 'rb') as file:
            part = MIMEApplication(file.read(), Name=os.path.basename(attachment_file))
        
        # Add header
        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(attachment_file)}"'
        msg.attach(part)
        
        # Connect to SMTP server and send email
        logger.info(f"Connecting to SMTP server {smtp_server}:{smtp_port}")
        server = smtplib.SMTP(smtp_server, smtp_port)
        server.set_debuglevel(0)
        
        if use_tls:
            logger.info("Starting TLS connection")
            server.starttls()
            
        if use_auth:
            logger.info(f"Logging in as {smtp_username}")
            server.login(smtp_username, smtp_password)
            
        logger.info(f"Sending email to {len(recipients)} recipients")
        server.sendmail(from_email, recipients, msg.as_string())
        server.quit()
        
        logger.info("Email sent successfully")
        
    except Exception as e:
        logger.error(f"Error sending email: {str(e)}")
        raise

def validate_config(config):
    """
    Validate the configuration and check required fields.
    
    Args:
        config: Configuration dictionary
        
    Returns:
        bool: True if configuration is valid
    """
    required_fields = ['query_file', 'csv_output_dir', 'smtp_server', 'from_email', 'recipients']
    
    # Database-specific required fields
    db_type = config.get('db_type', 'sqlite').lower()
    
    if db_type == 'sqlite':
        required_fields.append('db_connection')
    elif db_type == 'oracle':
        required_fields.extend(['db_connection', 'db_user', 'db_password'])
    elif db_type == 'mssql':
        if config.get('db_connection'):
            # Using direct connection string, no other fields required
            pass
        elif config.get('use_windows_auth', '').lower() == 'true':
            required_fields.extend(['db_server', 'db_name'])
        else:
            required_fields.extend(['db_server', 'db_name', 'db_user', 'db_password'])
    
    missing_fields = [field for field in required_fields if field not in config]
    
    if missing_fields:
        logger.error(f"Missing required configuration fields: {', '.join(missing_fields)}")
        return False
        
    return True

def main():
    """Main function to orchestrate the process."""
    if len(sys.argv) != 2:
        logger.error("Usage: python email_sql_csv_report.py <config_file>")
        sys.exit(1)
        
    config_file = sys.argv[1]
    
    try:
        # Read configuration
        logger.info(f"Reading configuration from {config_file}")
        config = read_config(config_file, required_fields=[])
        
        # Validate configuration
        if not validate_config(config):
            sys.exit(1)
            
        # Read SQL query from file
        query_file = config.get('query_file')
        logger.info(f"Reading SQL query from {query_file}")
        with open(query_file, 'r') as f:
            query = f.read()
            
        # Connect to database
        connection = connect_to_database(config, logger)
        
        try:
            # Run query
            logger.info("Executing SQL query")
            df = run_query(connection, query, logger)
            
            # Update the configuration to use a directory for CSV output
            csv_output_dir = config.get('csv_output_dir', '.')

            # Ensure the CSV output directory exists
            os.makedirs(csv_output_dir, exist_ok=True)

            # Generate a timestamped filename
            csv_filename = f"query_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            csv_output = os.path.join(csv_output_dir, csv_filename)

            # Export to CSV
            logger.info(f"Exporting results to {csv_output}")
            csv_file = export_to_csv(df, csv_output, logger)
            
            # Send email with attachment
            logger.info("Sending email with CSV attachment")
            send_email(config, csv_file)
            
            logger.info("Process completed successfully")
            
        finally:
            # Close database connection
            connection.close()
            logger.info("Database connection closed")
            
    except Exception as e:
        logger.error(f"Error in main process: {str(e)}")
        logger.error(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    main() 