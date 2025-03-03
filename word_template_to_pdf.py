#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Word Template to PDF Converter

This script replaces bracketed fields in a Word document with values from an Excel file,
then converts the result to PDF format.
For example, it will replace [First Name] or [First_Name] with "John" based on Excel data.

Requirements:
    - Python 3.6 or higher
    - Required Python packages:
        pip install python-docx==0.8.11    # For Word document handling
        pip install openpyxl==3.0.10       # For Excel file handling
        pip install docx2pdf==0.1.8        # For Word to PDF conversion on Windows
        pip install appscript==1.2.2       # For Word to PDF conversion on macOS
    - Microsoft Word must be installed (Windows or macOS)
    - If Microsoft Word is not available on Unix systems:
        macOS: brew install libreoffice
        Linux: sudo apt-get install libreoffice

Usage:
    1. Prepare your files:
       - Word template: Use bracketed fields like [Field Name] where you want replacements
       - Excel file: First row should contain headers matching the bracketed field names
         (without the brackets). For example, if [First Name] is in the Word doc,
         the Excel should have a column header "First Name"
       - Config file: Text file with the following format:
            excel_file = path/to/data.xlsx
            template = path/to/template.docx
            output_directory = path/to/output
            filename_field1 = First Name  # Optional - uses timestamp if both fields omitted
            filename_field2 = Last Name   # Optional - uses timestamp if both fields omitted
            keep_word_file = false        # Optional - set to true to keep both .docx and .pdf
    
    2. Run the script:
       python word_template_to_pdf.py <config_file>
    
       Example:
       python word_template_to_pdf.py config.txt

Note:
    - Field names in Word doc must match Excel headers exactly (excluding brackets)
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
import platform

# Import appropriate conversion module based on platform
if platform.system() == 'Darwin':  # macOS
    try:
        from appscript import app, k, mactypes
        WORD_APP = None  # Will be initialized on first use
    except ImportError:
        from docx2pdf import convert
else:
    from docx2pdf import convert

def normalize_field_name(name):
    """
    Normalize field names by converting spaces to underscores and vice versa.
    This allows matching both [First Name] and [First_Name] formats.
    
    Args:
        name (str): Field name to normalize
        
    Returns:
        list: List of possible field name variations
    """
    # Remove any leading/trailing whitespace
    name = name.strip()
    
    # Create variations
    with_spaces = name.replace('_', ' ')
    with_underscores = name.replace(' ', '_')
    
    # Return unique variations
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

def convert_to_pdf(docx_path, pdf_path):
    """
    Convert Word document to PDF using docx2pdf.
    
    Args:
        docx_path (str): Path to the Word document
        pdf_path (str): Path where the PDF should be saved
        
    Returns:
        bool: True if conversion was successful, False otherwise
    """
    try:
        # Use docx2pdf for all platforms
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        
        # Verify the PDF was created
        if not os.path.exists(pdf_path):
            raise Exception("PDF file was not created")
        if os.path.getsize(pdf_path) == 0:
            raise Exception("PDF file is empty")
            
        return True
            
    except Exception as e:
        print(f"Error converting to PDF: {e}")
        return False

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
    
    # Remove leading/trailing spaces and dots
    filename = filename.strip('. ')
    
    # Replace multiple spaces/underscores with single ones
    filename = re.sub(r'[\s_]+', '_', filename)
    
    # Limit length (Windows has a 255 character limit)
    max_length = 200  # Leave room for path and extension
    if len(filename) > max_length:
        filename = filename[:max_length]
    
    return filename

def main():
    """Main function to process Word documents."""
    if len(sys.argv) != 2:
        print("Usage: python word_template_to_pdf.py <config_file>")
        sys.exit(1)
    
    try:
        total_start_time = time.time()
        config = read_config(sys.argv[1])
        
        # Extract configuration
        excel_file = config['excel_file']
        word_template = config['template']
        output_directory = config['output_directory']
        filename_field1 = config.get('filename_field1', '')
        filename_field2 = config.get('filename_field2', '')
        keep_word = config.get('keep_word_file', '').lower() == 'true'
        
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
        
        # Verify all template fields exist in Excel headers (checking all variations)
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
            # Use the header's actual format
            filename_field1 = next(var for var in variations1 if var in headers)
            
        if filename_field2:
            variations2 = normalize_field_name(filename_field2)
            if not any(var in headers for var in variations2):
                raise ValueError(f"Specified filename field '{filename_field2}' not found in Excel headers")
            # Use the header's actual format
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
                
                # Create temporary Word document
                temp_doc = Document(word_template)
                replace_fields_in_document(temp_doc, data)
                
                # Save Word document if keeping it or as temporary file
                docx_path = os.path.join(output_directory, f"{filename}.docx")
                pdf_path = os.path.join(output_directory, f"{filename}.pdf")
                
                # Handle duplicate filenames
                counter = 1
                while os.path.exists(pdf_path) or os.path.exists(docx_path):
                    filename = f"{filename}_{counter}"
                    docx_path = os.path.join(output_directory, f"{filename}.docx")
                    pdf_path = os.path.join(output_directory, f"{filename}.pdf")
                    counter += 1
                
                # Save the Word document
                temp_doc.save(docx_path)
                
                # Convert to PDF
                if convert_to_pdf(docx_path, pdf_path):
                    success_count += 1
                    elapsed_time = time.time() - start_time
                    print(f"Processed {processed_count}/{total_files}: {filename}.pdf in {elapsed_time:.1f} seconds")
                    
                    # Remove Word document if not keeping it
                    if not keep_word and os.path.exists(docx_path):
                        os.remove(docx_path)
                else:
                    print(f"Failed to convert to PDF: {filename}")
                    
            except Exception as e:
                print(f"Error processing row {processed_count}: {str(e)}")
                traceback.print_exc()
        
        # Print summary
        total_time = time.time() - total_start_time
        print("\nProcessing Summary:")
        print(f"Total files processed: {success_count}/{total_files}")
        print(f"Total processing time: {total_time:.1f} seconds")
        print(f"Average time per file: {(total_time/total_files):.1f} seconds")
        print(f"Output directory: {os.path.abspath(output_directory)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 