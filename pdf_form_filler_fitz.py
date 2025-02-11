#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF Form Filler (PyMuPDF version)

This script fills PDF forms with data from an Excel file and optionally flattens specified fields.
All processing is done locally to maintain data privacy. It creates one filled PDF for each row 
in your Excel file.

Requirements:
    - Python 3.6 or higher
    - Required packages (install exact versions for compatibility):
        pip install openpyxl==3.0.10     # For Excel file handling
        pip install PyMuPDF==1.21.1      # For PDF form filling and flattening

    Installation steps:
    1. First, uninstall any existing versions:
        pip uninstall openpyxl PyMuPDF
    
    2. Then install the exact versions:
        pip install openpyxl==3.0.10 PyMuPDF==1.21.1

Usage:
    1. Prepare your files:
       - Excel file: First row should contain headers that match PDF form field names
       - PDF template: Should have fillable form fields
       - Fields config file: Text file with one field name per line to be flattened
    
    2. Run the script:
       python pdf_form_filler_fitz.py <excel_file> <pdf_template> <output_directory> [fields_to_flatten.txt]
    
       Example:
       python pdf_form_filler_fitz.py data.xlsx template.pdf output_forms fields.txt
    
    3. The script will:
       - Create one filled PDF for each row in your Excel file
       - Flatten the specified form fields (if fields config file is provided)
       - Save the PDFs in the specified output directory
       - Name each file with a timestamp and sequence number

Important Notes:
    1. The Excel column headers must match the PDF form field names exactly
    2. All processing is done locally - no data is sent over the internet
    3. Make sure you have write permissions in the output directory
    4. Back up your PDF template and Excel file before running the script
    5. Flattened fields cannot be edited after processing

Troubleshooting:
    1. Check that your Excel headers match the PDF form field names
    2. Ensure your PDF template has fillable form fields
    3. Verify that all required files exist and are accessible
    4. Check that the fields listed in your flatten config file exist in the PDF
    5. If you get PDF errors, make sure you have the correct package versions

License:
    This script is provided as-is under the MIT License.
"""

import sys
import os
from datetime import datetime
from openpyxl import load_workbook
import fitz

def read_excel_data(excel_path):
    """Read data from Excel file."""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        sheet = wb.active
        
        # Get headers from first row
        headers = [cell.value for cell in sheet[1]]
        
        # Read all rows (excluding header)
        data = []
        for row in list(sheet.rows)[1:]:
            row_data = {headers[i]: cell.value 
                       for i, cell in enumerate(row) 
                       if i < len(headers)}
            data.append(row_data)
            
        return headers, data
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None, None

def read_fields_to_flatten(fields_file):
    """Read list of fields to flatten from file."""
    try:
        if not fields_file:
            return set()
        with open(fields_file, 'r') as f:
            return {line.strip() for line in f if line.strip()}
    except Exception as e:
        print(f"Error reading fields config file: {e}")
        return set()

def fill_pdf_form(template_path, data_row, output_path, fields_to_flatten):
    """Fill a single PDF form and flatten specified fields."""
    try:
        # Open the PDF template
        doc = fitz.open(template_path)
        
        # Get form fields
        form_fields = doc.get_form_text_fields()
        if not form_fields:
            print("Warning: No form fields found in PDF")
        else:
            print(f"Found {len(form_fields)} form fields in PDF")
            
            # Fill form fields
            for field_name, value in data_row.items():
                if field_name in form_fields:
                    # Handle empty values
                    str_value = str(value).strip() if value is not None else ''
                    doc.set_form_text_field(field_name, str_value)
                    
                    # Flatten this field if it's in the list
                    if field_name in fields_to_flatten:
                        widgets = doc.get_form_text_widgets(field_name)
                        for widget in widgets:
                            # Get the field's current value and appearance
                            field_value = widget.field_value
                            if field_value:
                                # Create text annotation at the field's location
                                rect = widget.rect
                                page = doc[widget.page_number]
                                page.insert_text(rect.br, field_value)
                                # Remove the form field
                                widget.reset()
                else:
                    print(f"Warning: Field '{field_name}' not found in PDF form")
        
        # Save the filled PDF
        doc.save(output_path)
        doc.close()
        return True
        
    except Exception as e:
        print(f"Error filling PDF form: {e}")
        return False

def main():
    if len(sys.argv) not in [4, 5]:
        print("Usage: python pdf_form_filler_fitz.py <excel_file> <pdf_template> <output_directory> [fields_to_flatten.txt]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    pdf_template = sys.argv[2]
    output_dir = sys.argv[3]
    fields_file = sys.argv[4] if len(sys.argv) == 5 else None
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
    # Read fields to flatten
    fields_to_flatten = read_fields_to_flatten(fields_file)
    if fields_file:
        print(f"\nFields to flatten: {', '.join(fields_to_flatten)}")
    
    print("\nReading Excel data...")
    headers, data = read_excel_data(excel_path)
    if not headers or not data:
        print("Failed to read Excel data")
        sys.exit(1)
    
    print(f"\nFound {len(data)} rows to process")
    print(f"Fields available: {', '.join(headers)}")
    
    # Process each row
    success_count = 0
    for i, row_data in enumerate(data, 1):
        print(f"\nProcessing row {i}/{len(data)}")
        
        # Generate output filename
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"filled_form_{timestamp}_{i}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        
        # Fill the form
        if fill_pdf_form(pdf_template, row_data, output_path, fields_to_flatten):
            print(f"Successfully created: {output_filename}")
            success_count += 1
        else:
            print(f"Failed to process row {i}")
    
    print(f"\nProcessing complete. Successfully filled {success_count} out of {len(data)} forms.")
    print(f"Output files are in: {output_dir}")

if __name__ == "__main__":
    main() 