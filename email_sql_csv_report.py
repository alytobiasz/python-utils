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
db_connection = mydatabase.db  # SQLite filename or connection string for other DBs
db_user = user
db_password = password
db_server = server_name  # For MS SQL Server
db_port = 1433  # For MS SQL Server (default)
db_name = database_name  # For MS SQL Server
db_driver = ODBC Driver 17 for SQL Server  # For MS SQL Server

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
import csv
import sqlite3
import pandas as pd
import smtplib
import traceback
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import logging

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

def read_config(config_file):
    """
    Read the configuration file and return a dictionary of settings.
    """
    config = {}
    try:
        with open(config_file, 'r') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                    
                if '=' in line:
                    key, value = line.split('=', 1)
                    config[key.strip()] = value.strip()
    except Exception as e:
        logger.error(f"Error reading configuration file: {str(e)}")
        raise
        
    return config

def connect_to_database(config):
    """
    Connect to the database based on the configuration.
    
    Returns:
        connection: Database connection object
    """
    db_type = config.get('db_type', 'sqlite').lower()
    db_connection = config.get('db_connection', '')
    
    try:
        if db_type == 'sqlite':
            connection = sqlite3.connect(db_connection)
            logger.info(f"Connected to SQLite database: {db_connection}")
            
        elif db_type == 'oracle':
            import oracledb
            connection = oracledb.connect(user=config.get('db_user', ''),
                                          password=config.get('db_password', ''),
                                          dsn=config.get('db_connection', ''))
            logger.info(f"Connected to Oracle database: {config.get('db_connection', '')}")
            
        elif db_type == 'mssql':
            import pyodbc
            # Build connection string
            server = config.get('db_server', '')
            port = config.get('db_port', '1433')
            database = config.get('db_name', '')
            username = config.get('db_user', '')
            password = config.get('db_password', '')
            driver = config.get('db_driver', 'ODBC Driver 17 for SQL Server')
            
            conn_str = f"DRIVER={{{driver}}};SERVER={server},{port};DATABASE={database};UID={username};PWD={password}"
            
            # Use provided connection string if it exists
            if db_connection:
                conn_str = db_connection
                
            connection = pyodbc.connect(conn_str)
            logger.info(f"Connected to MS SQL Server database: {database} on {server}")
            
        else:
            raise ValueError(f"Unsupported database type: {db_type}")
            
        return connection
        
    except Exception as e:
        logger.error(f"Error connecting to {db_type} database: {str(e)}")
        raise

def run_query(connection, query):
    """
    Run the SQL query and return the results as a pandas DataFrame.
    
    Args:
        connection: Database connection object
        query: SQL query string
        
    Returns:
        pd.DataFrame: Results of the query
    """
    try:
        df = pd.read_sql_query(query, connection)
        logger.info(f"Query executed successfully. Retrieved {len(df)} rows.")
        return df
    except Exception as e:
        logger.error(f"Error executing query: {str(e)}")
        raise

def export_to_csv(df, output_file):
    """
    Export the DataFrame to a CSV file.
    
    Args:
        df: pandas DataFrame
        output_file: Path to save the CSV file
        
    Returns:
        str: Path to the saved CSV file
    """
    try:
        df.to_csv(output_file, index=False, quoting=csv.QUOTE_MINIMAL)
        logger.info(f"Data exported to CSV file: {output_file}")
        return output_file
    except Exception as e:
        logger.error(f"Error exporting to CSV: {str(e)}")
        raise

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
    required_fields = [
        'db_connection',
        'query_file',
        'csv_output_dir',
        'smtp_server',
        'from_email',
        'recipients'
    ]
    
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
        config = read_config(config_file)
        
        # Validate configuration
        if not validate_config(config):
            sys.exit(1)
            
        # Read SQL query from file
        query_file = config.get('query_file')
        logger.info(f"Reading SQL query from {query_file}")
        with open(query_file, 'r') as f:
            query = f.read()
            
        # Connect to database
        connection = connect_to_database(config)
        
        try:
            # Run query
            logger.info("Executing SQL query")
            df = run_query(connection, query)
            
            # Update the configuration to use a directory for CSV output
            csv_output_dir = config.get('csv_output_dir', '.')

            # Ensure the CSV output directory exists
            os.makedirs(csv_output_dir, exist_ok=True)

            # Generate a timestamped filename
            csv_filename = f"query_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
            csv_output = os.path.join(csv_output_dir, csv_filename)

            # Export to CSV
            logger.info(f"Exporting results to {csv_output}")
            csv_file = export_to_csv(df, csv_output)
            
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