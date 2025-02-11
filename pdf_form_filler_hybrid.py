#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
PDF Form Filler (Hybrid version)

This script uses pdfrw to fill PDF forms and PyMuPDF (fitz) to flatten specified fields.
All processing is done locally to maintain data privacy.

Requirements:
    - Python 3.6 or higher
    - Required packages (install exact versions for compatibility):
        pip install openpyxl==3.0.10     # For Excel file handling
        pip install pdfrw==0.4.0         # For PDF form filling
        pip install PyMuPDF==1.21.1      # For field flattening

    Installation steps:
    1. First, uninstall any existing versions:
        pip uninstall openpyxl pdfrw PyMuPDF
    
    2. Then install the exact versions:
        pip install openpyxl==3.0.10 pdfrw==0.4.0 PyMuPDF==1.21.1

Usage:
    1. Prepare your files:
       - Excel file: First row should contain headers that match PDF form field names
       - PDF template: Should have fillable form fields
       - Fields config file: Text file with one field name per line to be flattened
    
    2. Run the script:
       python pdf_form_filler_hybrid.py <excel_file> <pdf_template> <output_directory> [fields_to_flatten.txt]
    
       Example:
       python pdf_form_filler_hybrid.py data.xlsx template.pdf output_forms fields.txt

Important Notes:
    1. The Excel column headers must match the PDF form field names exactly
    2. All processing is done locally - no data is sent over the internet
    3. Make sure you have write permissions in the output directory
    4. Back up your PDF template and Excel file before running the script
    5. Flattened fields cannot be edited after processing
"""

import sys
import os
from datetime import datetime
import time
from openpyxl import load_workbook
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfObject
import fitz

def read_excel_data(excel_path):
    """Read data from Excel file."""
    try:
        wb = load_workbook(excel_path, read_only=True, data_only=True)
        sheet = wb.active
        headers = [cell.value for cell in sheet[1]]
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

def fill_pdf_form(template_path, data_row, temp_output_path):
    """Fill PDF form using pdfrw."""
    try:
        template = PdfReader(template_path)
        writer = PdfWriter()
        
        if not template.Root.AcroForm:
            return False
            
        form_fields = template.Root.AcroForm.Fields
        if not form_fields:
            return False
        
        for field in form_fields:
            if field.T:
                key = field.T[1:-1]
                if key in data_row:
                    value = data_row[key]
                    if value is None or str(value).strip() == '':
                        field.V = ''
                    else:
                        field.V = str(value).strip()
                    field.AP = ''
        
        template.Root.AcroForm.update(PdfDict(
            NeedAppearances=PdfObject('true')
        ))
        
        for page in template.pages:
            writer.addpage(page)
        
        writer.write(temp_output_path)
        return True
        
    except Exception:
        return False

def flatten_fields(input_path, output_path, fields_to_flatten):
    """Flatten specified fields using PyMuPDF."""
    doc = None
    try:
        doc = fitz.open(input_path)
        
        for page in doc:
            for field in page.widgets():
                try:
                    # Get field name directly from widget
                    field_name = field.field_name
                    if not field_name or field_name not in fields_to_flatten:
                        continue
                        
                    # Get field value directly
                    field_value = field.field_value
                    if not field_value or not str(field_value).strip():
                        # Just remove empty fields
                        page.delete_widget(field)
                        continue
                    
                    # Get field position and properties
                    rect = field.rect
                    font_size = 12  # Use consistent font size
                    
                    # Insert text first
                    page.insert_text(
                        point=(rect.x0 + 2, rect.y0 + font_size),
                        text=str(field_value),
                        fontsize=font_size,
                        color=(0, 0, 0)
                    )
                    
                    # Then remove the widget
                    page.delete_widget(field)
                    
                except:
                    # If any error occurs with a field, try to just remove it
                    try:
                        page.delete_widget(field)
                    except:
                        pass
        
        # Save with optimization
        doc.save(output_path, garbage=4, deflate=True, clean=True)
        doc.close()
        return os.path.exists(output_path) and os.path.getsize(output_path) > 0
            
    except Exception:
        return False
    finally:
        if doc:
            try:
                doc.close()
            except:
                pass

def process_pdf(template_path, data_row, output_path, fields_to_flatten):
    """Process a single PDF form - fill and flatten."""
    try:
        temp_path = output_path + '.temp.pdf'
        
        if not fill_pdf_form(template_path, data_row, temp_path):
            return False
        
        if fields_to_flatten:
            if not flatten_fields(temp_path, output_path, fields_to_flatten):
                return False
            try:
                os.remove(temp_path)
            except:
                pass
        else:
            try:
                os.replace(temp_path, output_path)
            except:
                return False
        
        return True
    except Exception:
        return False

def main():
    if len(sys.argv) not in [4, 5]:
        print("Usage: python pdf_form_filler_hybrid.py <excel_file> <pdf_template> <output_directory> [fields_to_flatten.txt]")
        sys.exit(1)
    
    excel_path = sys.argv[1]
    pdf_template = sys.argv[2]
    output_dir = sys.argv[3]
    fields_file = sys.argv[4] if len(sys.argv) == 5 else None
    
    os.makedirs(output_dir, exist_ok=True)
    
    fields_to_flatten = read_fields_to_flatten(fields_file)
    
    print("\nReading Excel data...")
    headers, data = read_excel_data(excel_path)
    if not headers or not data:
        print("Failed to read Excel data")
        sys.exit(1)
    
    total_start_time = time.time()
    success_count = 0
    
    print(f"\nProcessing {len(data)} files...")
    
    for i, row_data in enumerate(data, 1):
        start_time = time.time()
        
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_filename = f"filled_form_{timestamp}_{i}.pdf"
        output_path = os.path.join(output_dir, output_filename)
        
        if process_pdf(pdf_template, row_data, output_path, fields_to_flatten):
            elapsed = time.time() - start_time
            print(f"Processed file {i}/{len(data)} in {elapsed:.1f} seconds")
            success_count += 1
        else:
            print(f"Failed to process file {i}/{len(data)}")
    
    total_time = time.time() - total_start_time
    print(f"\nProcessing complete:")
    print(f"Successfully processed {success_count} out of {len(data)} files")
    print(f"Total processing time: {total_time:.1f} seconds")
    print(f"Average time per file: {(total_time/len(data)):.1f} seconds")
    print(f"Output files are in: {output_dir}")

if __name__ == "__main__":
    main() 