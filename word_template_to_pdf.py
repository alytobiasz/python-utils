#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Word Template to PDF Converter

This script combines docx_template_filler.py and docx_to_pdf.py to:
1. Fill Word templates with data from Excel
2. Convert the resulting documents to PDF

Requirements:
    - All requirements from docx_template_filler.py
    - All requirements from docx_to_pdf.py
    - Microsoft Word must be installed (Windows or macOS)

Usage:
    python word_template_to_pdf.py <config_file>

Example:
    python word_template_to_pdf.py config.txt

The config file format is the same as docx_template_filler.py, with one additional option:
    keep_word_file = false  # Optional - set to true to keep both .docx and .pdf
"""

import sys
import os
import platform
import time
from datetime import datetime

# Import the functions from both scripts
from docx_template_filler import read_config, fill_docx_templates
from docx_to_pdf import create_pdfs

def main():
    """Main function to coordinate template filling and PDF conversion."""
    if len(sys.argv) != 2:
        print("Usage: python word_template_to_pdf.py <config_file>")
        sys.exit(1)
    
    # Check if running on supported OS
    if platform.system() not in ['Windows', 'Darwin']:
        print("Error: This script only supports Windows and macOS")
        sys.exit(1)
    
    try:
        total_start_time = time.time()
        
        # Read configuration
        config = read_config(sys.argv[1])
        base_output_dir = config['output_directory']
        keep_word = config.get('keep_word_file', '').lower() == 'true'
        
        # Create timestamped directory if base directory exists and is not empty
        if os.path.exists(base_output_dir) and os.listdir(base_output_dir):
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_dir = os.path.join(base_output_dir, f'batch_{timestamp}')
            print(f"\nOutput directory not empty, creating new directory: {output_dir}")
        else:
            output_dir = base_output_dir
            
        # Create output directory
        os.makedirs(output_dir, exist_ok=True)
        
        # Update config with new output directory
        config['output_directory'] = output_dir
        
        # Step 1: Fill templates to create Word documents
        print("\nStep 1: Generating Word documents...")
        docx_success, docx_total = fill_docx_templates(config)
        
        # Step 2: Convert Word documents to PDF
        print("\nStep 2: Converting to PDF...")
        pdf_success, pdf_total = create_pdfs(output_dir, output_dir)
        
        # Clean up Word files if not keeping them
        if not keep_word:
            print("\nCleaning up Word files...")
            for file in os.listdir(output_dir):
                if file.endswith('.docx'):
                    try:
                        os.remove(os.path.join(output_dir, file))
                    except Exception as e:
                        print(f"Warning: Could not remove {file}: {e}")
        
        # Print completion message
        total_time = time.time() - total_start_time
        print(f"\nProcessing completed in {total_time:.1f} seconds")
        print(f"Word documents generated: {docx_success}/{docx_total}")
        print(f"PDF files created: {pdf_success}/{pdf_total}")
        print(f"Output directory: {os.path.abspath(output_dir)}")
        if keep_word:
            print(f"Word files directory: {os.path.abspath(output_dir)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 