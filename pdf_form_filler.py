#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF Form Filler

This script fills PDF forms with data from an Excel file. All processing is done locally 
to maintain data privacy. It creates one filled PDF for each row in your Excel file.

Requirements:
    - Python 3.6 or higher
    - Required packages (install using pip):
        pip install openpyxl pdfrw

Setup:
    1. Create a virtual environment (recommended):
       python -m venv .venv
       source .venv/bin/activate  # On Windows: .venv\Scripts\activate
    
    2. Install dependencies:
       pip install openpyxl pdfrw

Usage:
    1. Prepare your Excel file:
       - First row should contain headers that match your PDF form field names exactly
       - Each subsequent row contains the data to fill in one PDF form
    
    2. Run the script:
       python pdf_form_filler.py <excel_file> <pdf_template> <output_directory>
    
       Example:
       python pdf_form_filler.py data.xlsx template.pdf output_forms
    
    3. The script will:
       - Create one filled PDF for each row in your Excel file
       - Save the PDFs in the specified output directory
       - Name each file with a timestamp and sequence number

Important Notes:
    1. The Excel column headers must match the PDF form field names exactly
    2. All processing is done locally - no data is sent over the internet
    3. Make sure you have write permissions in the output directory
    4. Back up your PDF template and Excel file before running the script

Troubleshooting:
    1. Check that your Excel headers match the PDF form field names
    2. Ensure your PDF template has fillable form fields
    3. Verify that all required files exist and are accessible
    4. Check that you have write permissions in the output directory

License:
    This script is provided as-is under the MIT License.
"""

import sys
import os
from datetime import datetime
from openpyxl import load_workbook
from pdfrw import PdfReader, PdfWriter, PdfDict

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

def fill_pdf_form(template_path, data_row, output_path):
    """Fill a single PDF form with data."""
    try:
        # Read the PDF template
        template = PdfReader(template_path)
        
        # Ensure we have form fields to work with
        if not template.Root.AcroForm:
            print("Error: No form fields found in PDF template")
            return False
            
        # Set NeedAppearances flag first
        template.Root.AcroForm.update(PdfDict(NeedAppearances=True))
        
        # Fill in the form fields
        for field_name, value in data_row.items():
            if template.Root.AcroForm.Fields:
                for field in template.Root.AcroForm.Fields:
                    if field.T and field.T[1:-1] == field_name:
                        # Handle empty values (None, empty string, or whitespace)
                        if value is None or str(value).strip() == '':
                            field.V = ''
                        else:
                            field.V = str(value).strip()
                        
                        # Clear any existing appearance streams
                        field.AP = ''
        
        # Write the filled PDF
        writer = PdfWriter()
        writer.write(output_path, template)
        return True
    except Exception as e:
        print(f"Error filling PDF form: {e}")
        return False

def main():
    if len(sys.argv) != 4:
        print("Usage: python pdf_form_filler.py <excel_file> <pdf_template> <output_directory>")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    pdf_template = sys.argv[2]
    output_dir = sys.argv[3]
    
    # Create output directory if it doesn't exist
    os.makedirs(output_dir, exist_ok=True)
    
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
        if fill_pdf_form(pdf_template, row_data, output_path):
            print(f"Successfully created: {output_filename}")
            success_count += 1
        else:
            print(f"Failed to process row {i}")
    
    print(f"\nProcessing complete. Successfully filled {success_count} out of {len(data)} forms.")
    print(f"Output files are in: {output_dir}")

if __name__ == "__main__":
    main() 