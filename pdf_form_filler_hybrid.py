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
from openpyxl import load_workbook
from pdfrw import PdfReader, PdfWriter, PdfDict, PdfName, PdfObject
import fitz
import pymupdf
import time

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

def fill_pdf_form(template_path, data_row, temp_output_path):
    """Fill PDF form using pdfrw."""
    try:
        # Read the PDF template
        template = PdfReader(template_path)
        
        # Create PDF writer
        writer = PdfWriter()
        
        # Check if PDF has form fields
        if not template.Root.AcroForm:
            print("Error: No form fields found in PDF template")
            return False
            
        # Get form fields
        form_fields = template.Root.AcroForm.Fields
        if not form_fields:
            print("Warning: No form fields found in PDF")
        
        # Update form fields
        fields_filled = 0
        for field in form_fields:
            if field.T:
                key = field.T[1:-1]  # Remove parentheses from field name
                if key in data_row:
                    value = data_row[key]
                    if value is None or str(value).strip() == '':
                        field.V = ''
                    else:
                        field.V = str(value).strip()
                    field.AP = ''
                    fields_filled += 1
                
        # Set form flags
        template.Root.AcroForm.update(PdfDict(
            NeedAppearances=PdfObject('true')
        ))
        
        # Add all pages to the writer
        for page in template.pages:
            writer.addpage(page)
        
        # Write the filled PDF
        writer.write(temp_output_path)
        return True
        
    except Exception as e:
        print(f"Error filling PDF form: {e}")
        return False

def flatten_fields(input_path, output_path, fields_to_flatten):
    """Flatten specified fields using PyMuPDF."""
    pymupdf.TOOLS.mupdf_display_errors(False)
    doc = None
    try:
        # Open the filled PDF
        doc = fitz.open(input_path)
        
        # Process each page
        for page_num, page in enumerate(doc):
            try:
                # Get form fields (widgets) on the page
                fields = list(page.widgets())
                
                for field in fields:
                    try:
                        # Get field name safely
                        field_name = getattr(field, 'field_name', None)
                        if not field_name:
                            continue
                            
                        if field_name in fields_to_flatten:
                            
                            # Get field value safely
                            try:
                                field_value = field.field_value
                            except:
                                try:
                                    # Alternative way to get value
                                    field_value = field.text
                                except:
                                    print(f"Could not get value for field {field_name}")
                                    field_value = None
                            
                            # Check if field has a value
                            if field_value and str(field_value).strip():
                                try:
                                    # Get field rectangle
                                    rect = field.rect
                                    
                                    # Try to get font size from field
                                    try:
                                        font_size = field.font_size
                                        if not font_size or font_size <= 0:
                                            font_size = 12
                                    except:
                                        font_size = 12
                                                                        
                                    # Calculate text position (slightly offset from top-left)
                                    x = rect.x0 + 2  # 2 point offset from left
                                    y = rect.y0 + font_size  # offset by font size from top
                                    
                                    # Insert text at the calculated position
                                    page.insert_text(
                                        point=(x, y),
                                        text=str(field_value),
                                        fontsize=font_size,
                                        color=(0, 0, 0)  # Black text
                                    )
                                    
                                    # Try to remove the form field
                                    try:
                                        page.delete_widget(field)
                                    except Exception as e:
                                        print(f"Warning: Could not delete widget for {field_name}: {e}")
                                        
                                except Exception as e:
                                    print(f"Error processing field '{field_name}': {e}")
                                    import traceback
                                    print(traceback.format_exc())
                                    continue
                            else:
                                # Field is empty, just remove the widget
                                try:
                                    page.delete_widget(field)
                                except Exception as e:
                                    print(f"Warning: Could not remove empty field {field_name}: {e}")
                                    
                    except Exception as field_error:
                        print(f"Error processing field: {field_error}")
                        import traceback
                        print(traceback.format_exc())
                        continue
                        
            except Exception as page_error:
                print(f"Error processing page {page_num + 1}: {page_error}")
                import traceback
                print(traceback.format_exc())
                continue
        
        # Save the modified PDF
        doc.save(output_path, garbage=4, deflate=True, clean=True)
        
        # Verify the file was saved
        if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
            return True
        else:
            print("Error: Failed to save PDF or file is empty")
            return False
            
    except Exception as e:
        print(f"Error flattening PDF: {e}")
        import traceback
        print("Traceback:")
        print(traceback.format_exc())
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
        # Create temporary file for intermediate step
        temp_path = output_path + '.temp.pdf'
        
        # Step 1: Fill the form using pdfrw
        if not fill_pdf_form(template_path, data_row, temp_path):
            print("Failed to fill PDF form")
            return False
        
        # Step 2: Flatten specified fields using PyMuPDF
        if fields_to_flatten:
            if not flatten_fields(temp_path, output_path, fields_to_flatten):
                print("Failed to flatten fields")
                return False
            # Remove temporary file
            try:
                os.remove(temp_path)
            except:
                pass
        else:
            # If no fields to flatten, just rename the temp file
            try:
                os.replace(temp_path, output_path)
            except Exception as e:
                print(f"Error moving file: {e}")
                return False
        
        return True
    except Exception as e:
        print(f"Error processing PDF: {e}")
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
            print(f"File {i}/{len(data)} - Success - {elapsed:.1f} seconds")
            success_count += 1
        else:
            elapsed = time.time() - start_time
            print(f"File {i}/{len(data)} - Failed - {elapsed:.1f} seconds")
    
    total_time = time.time() - total_start_time
    avg_time = total_time / len(data)
    
    print(f"\nProcessing complete:")
    print(f"Successfully processed: {success_count}/{len(data)} files")
    print(f"Total time: {total_time:.1f} seconds")
    print(f"Average time per file: {avg_time:.1f} seconds")
    print(f"Output directory: {output_dir}")

if __name__ == "__main__":
    main() 