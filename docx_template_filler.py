#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX Template Filler

This script replaces bracketed fields in a Word (.docx) document with values from an Excel file.
For example, it will replace [First Name] or [First_Name] with "John" based on Excel data.

Requirements:
    - Python 3.6 or higher
    - Required Python packages:
        pip install python-docx==0.8.11    # For Word document handling
        pip install openpyxl==3.0.10       # For Excel file handling

Usage:
    1. Prepare your files:
       - DOCX template: Use bracketed fields like [Field Name] where you want replacements
       - Excel file: First row should contain headers matching the bracketed field names
         (without the brackets). For example, if [First Name] is in the document,
         the Excel should have a column header "First Name"
       - Config file: Text file with the following format:
            excel_file = path/to/data.xlsx
            template = path/to/template.docx
            output_directory = path/to/output
            filename_field1 = First Name  # Optional - uses timestamp if both fields omitted
            filename_field2 = Last Name   # Optional - uses timestamp if both fields omitted

    2. Run the script:
       python docx_template_filler.py <config_file>
    
       Example:
       python docx_template_filler.py config.txt

Note:
    - Field names in the document must match Excel headers exactly (excluding brackets)
    - Fields are case-sensitive: [First_Name] ≠ [first_name]
    - Output files will be named using the specified fields (or timestamp if omitted)
"""

import sys
import os
import re
from datetime import datetime
import time
from docx import Document
from openpyxl import load_workbook
import traceback

def normalize_field_name(name):
    """
    Normalize field names by converting spaces to underscores and vice versa.
    This allows matching both [First Name] and [First_Name] formats.
    
    Args:
        name (str): Field name to normalize
        
    Returns:
        list: List of possible field name variations
    """
    name = name.strip()
    with_spaces = name.replace('_', ' ')
    with_underscores = name.replace(' ', '_')
    return list(set([name, with_spaces, with_underscores]))

def find_fields_in_document(doc):
    """
    Find all bracketed fields in the Word document.
    
    Args:
        doc: Word document object
        
    Returns:
        set: Set of unique field names found (without brackets)
    """
    fields = set()
    pattern = r'\[([^\]]+)\]'
    
    # Search in paragraphs
    for paragraph in doc.paragraphs:
        matches = re.finditer(pattern, paragraph.text)
        fields.update(match.group(1) for match in matches)
    
    # Search in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                matches = re.finditer(pattern, cell.text)
                fields.update(match.group(1) for match in matches)
    
    return fields

def replace_fields_in_document(doc, data):
    """
    Replace all bracketed fields with corresponding values.
    
    Args:
        doc: Word document object
        data (dict): Dictionary of field names and their values
    """
    # Create a mapping of all possible field variations
    field_mapping = {}
    for key in data:
        variations = normalize_field_name(key)
        for variant in variations:
            field_mapping[variant] = data[key]
    
    # Replace in paragraphs
    for paragraph in doc.paragraphs:
        for field_name, value in field_mapping.items():
            if f"[{field_name}]" in paragraph.text:
                paragraph.text = paragraph.text.replace(f"[{field_name}]", 
                    str(value) if value is not None else '')
    
    # Replace in tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for field_name, value in field_mapping.items():
                    if f"[{field_name}]" in cell.text:
                        cell.text = cell.text.replace(f"[{field_name}]", 
                            str(value) if value is not None else '')

def read_config(config_path):
    """
    Read configuration from a file.
    
    Args:
        config_path (str): Path to the configuration file
        
    Returns:
        dict: Configuration parameters
    """
    config = {}
    required_fields = ['excel_file', 'template', 'output_directory']
    
    try:
        with open(config_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    key, value = map(str.strip, line.split('=', 1))
                    config[key] = value
    except Exception as e:
        raise ValueError(f"Error reading config file: {str(e)}")
    
    # Check for required fields
    missing = [field for field in required_fields if field not in config]
    if missing:
        raise ValueError(f"Missing required fields in config file: {', '.join(missing)}")
    
    return config

def sanitize_filename(filename):
    """
    Sanitize a filename by removing or replacing invalid characters.
    
    Args:
        filename (str): The filename to sanitize
        
    Returns:
        str: The sanitized filename
    """
    # Replace invalid characters with underscores
    invalid_chars = r'[<>:"/\\|?*]'
    filename = re.sub(invalid_chars, '_', filename)
    filename = filename.strip('. ')
    filename = re.sub(r'[\s_]+', '_', filename)
    
    # Limit length (Windows has a 255 character limit)
    max_length = 200
    if len(filename) > max_length:
        filename = filename[:max_length]
    
    return filename

def fill_docx_templates(config):
    """
    Fill Word document templates with data from Excel.
    
    Args:
        config (dict): Configuration dictionary containing:
            - excel_file: Path to Excel file with data
            - template: Path to Word template file
            - output_directory: Directory for output files
            - filename_field1: Optional field for filename generation
            - filename_field2: Optional field for filename generation
        
    Returns:
        tuple: (success_count, total_files) indicating number of successfully processed files
    """
    try:
        # Extract configuration
        excel_file = config['excel_file']
        word_template = config['template']
        output_directory = config['output_directory']
        filename_field1 = config.get('filename_field1', '')
        filename_field2 = config.get('filename_field2', '')
        
        # Create output directory if it doesn't exist
        os.makedirs(output_directory, exist_ok=True)
        
        # Load the template to find fields
        template_doc = Document(word_template)
        template_fields = find_fields_in_document(template_doc)
        print(f"\nFound {len(template_fields)} unique fields in Word template:")
        print(", ".join(sorted(template_fields)))
        
        # Read Excel data
        wb = load_workbook(filename=excel_file, data_only=True)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        
        # Verify all template fields exist in Excel headers
        missing_fields = []
        for field in template_fields:
            field_variations = normalize_field_name(field)
            if not any(var in headers for var in field_variations):
                missing_fields.append(field)
        
        if missing_fields:
            raise ValueError(f"Fields in Word template not found in Excel headers: {', '.join(missing_fields)}")
        
        # Verify filename fields exist in headers if specified
        if filename_field1:
            variations1 = normalize_field_name(filename_field1)
            if not any(var in headers for var in variations1):
                raise ValueError(f"Specified filename field '{filename_field1}' not found in Excel headers")
            filename_field1 = next(var for var in variations1 if var in headers)
            
        if filename_field2:
            variations2 = normalize_field_name(filename_field2)
            if not any(var in headers for var in variations2):
                raise ValueError(f"Specified filename field '{filename_field2}' not found in Excel headers")
            filename_field2 = next(var for var in variations2 if var in headers)
        
        if not filename_field1 and not filename_field2:
            print("No filename fields specified - using timestamps for output files")
        
        # Count total non-empty rows
        total_files = sum(1 for row in ws.iter_rows(min_row=2) if any(cell.value for cell in row))
        processed_count = 0
        success_count = 0
        
        # Process each row
        for row_cells in ws.iter_rows(min_row=2):
            row = [cell.value for cell in row_cells]
            if not any(row):  # Skip empty rows
                continue
            
            processed_count += 1
            start_time = time.time()
            
            try:
                # Create data dictionary
                data = {headers[i]: str(val) if val is not None else '' 
                       for i, val in enumerate(row)}
                
                # Generate output filename from specified fields
                if filename_field1 or filename_field2:
                    field1_value = data.get(filename_field1, '').strip()
                    field2_value = data.get(filename_field2, '').strip()
                    filename = f"{field1_value} {field2_value}".strip()
                else:
                    # Use timestamp if no fields specified
                    filename = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                # Sanitize filename
                filename = sanitize_filename(filename)
                
                # Create output path
                docx_path = os.path.join(output_directory, f"{filename}.docx")
                
                # Handle duplicate filenames
                counter = 1
                while os.path.exists(docx_path):
                    filename = f"{filename}_{counter}"
                    docx_path = os.path.join(output_directory, f"{filename}.docx")
                    counter += 1
                
                # Create and save the filled document
                doc = Document(word_template)
                replace_fields_in_document(doc, data)
                doc.save(docx_path)
                
                success_count += 1
                elapsed_time = time.time() - start_time
                print(f"Processed {processed_count}/{total_files}: {filename}.docx in {elapsed_time:.1f} seconds")
                
            except Exception as e:
                print(f"Error processing row {processed_count}: {str(e)}")
                traceback.print_exc()
        
        # Print summary
        print("\nProcessing Summary:")
        print(f"Total files processed: {success_count}/{total_files}")
        print(f"Output directory: {os.path.abspath(output_directory)}")
        
        return success_count, total_files
        
    except Exception as e:
        print(f"Error: {str(e)}")
        traceback.print_exc()
        sys.exit(1)

def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) != 2:
        print("Usage: python docx_template_filler.py <config_file>")
        sys.exit(1)
    
    config = read_config(sys.argv[1])
    fill_docx_templates(config)

if __name__ == "__main__":
    main() 