"""
SQL Query to CSV Report Generator

This script:
1. Connects to a database
2. Runs a SQL query
3. Exports the results to a CSV file

Requirements:
    - Python 3.6+
    - Required packages:
        - sqlite3 (built-in) or another DB driver depending on your database
        - pandas (for easy data handling and CSV export)
        - For MS SQL Server: pyodbc
        - For Oracle: oracledb

Usage:
    python create_sql_csv_report.py <config_file>

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
"""

import sys
import os
import logging
from datetime import datetime
import traceback

# Import utility functions
from utils import connect_to_database, run_query, export_to_csv, read_config

# Set up logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(f"logs/sql_csv_report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    ]
)
logger = logging.getLogger(__name__)

def validate_config(config):
    """
    Validate the configuration and check required fields.
    
    Args:
        config: Configuration dictionary
        
    Returns:
        bool: True if configuration is valid
    """
    required_fields = []
    
    # Common required fields
    required_fields.extend(['query_file', 'csv_output_dir'])
    
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
        logger.error("Usage: python create_sql_csv_report.py <config_file>")
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
            
            logger.info(f"Process completed successfully. CSV file saved to: {csv_file}")
            
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