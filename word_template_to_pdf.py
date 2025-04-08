#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Word Template to PDF Converter

This script combines docx_template_filler.py and docx_to_pdf.py to:
1. Fill Word templates with data from Excel
2. Convert the resulting documents to PDF

Requirements:
    - All requirements from docx_template_filler.py
    - All requirements from docx_to_pdf.py or libreoffice_docx_to_pdf.py depending on engine choice
    - Microsoft Word (Windows or macOS) or LibreOffice must be installed

Usage:
    python word_template_to_pdf.py <config_file>

Example:
    python word_template_to_pdf.py config.txt

The config file format is the same as docx_template_filler.py, with additional options:
    keep_word_file = false  # Optional - set to true to keep both .docx and .pdf
    conversion_engine = word  # Optional - 'word' or 'libreoffice' (default: 'word')
    max_threads = 1  # Optional - number of threads for PDF conversion (default: 1, recommended)
"""

import sys
import os
import platform
import time
import subprocess
from datetime import datetime

# Import the shared utility function for reading config files
from utils import read_config

# Import docx_to_pdf for conversion and dependency checking
try:
    from docx_to_pdf import check_dependencies, create_pdfs as create_pdfs_word
    dependencies_ok = check_dependencies()
except ImportError:
    # This will happen if running on an unsupported OS or missing dependencies
    dependencies_ok = False
    create_pdfs_word = None

# Import docx_template_filler as it doesn't depend on platform-specific libraries
from docx_template_filler import fill_docx_templates

# Try to import LibreOffice-related functions
libreoffice_module_available = False
libreoffice_installed = False
libreoffice_available = False

try:
    # Try to import both the conversion function and any available check functions
    from libreoffice_docx_to_pdf import (
        create_pdfs as create_pdfs_libreoffice,
        is_libreoffice_installed  # Try to import this function if it exists
    )
    libreoffice_module_available = True
    
    # Use the module's own check function if available
    libreoffice_installed = is_libreoffice_installed()
    libreoffice_available = libreoffice_installed
    
except ImportError:
    # Module not available at all
    libreoffice_module_available = False
    libreoffice_available = False

except AttributeError:
    # Module available but is_libreoffice_installed function doesn't exist
    libreoffice_module_available = True
    
    # Fall back to a simple check
    try:
        # Simple check: try to run the soffice command
        if platform.system() == 'Windows':
            result = subprocess.run(['where', 'soffice'], capture_output=True, text=True)
        else:  # macOS, Linux, etc.
            result = subprocess.run(['which', 'soffice'], capture_output=True, text=True)
        libreoffice_installed = result.returncode == 0
        libreoffice_available = libreoffice_installed
    except Exception:
        libreoffice_installed = False
        libreoffice_available = False

def main():
    """Main function to coordinate template filling and PDF conversion."""
    if len(sys.argv) != 2:
        print("Usage: python word_template_to_pdf.py <config_file>")
        sys.exit(1)
    
    try:
        total_start_time = time.time()
        
        # Read configuration
        config = read_config(sys.argv[1])
        base_output_dir = config['output_directory']
        keep_word = config.get('keep_word_file', '').lower() == 'true'
        
        # Get conversion engine preference (default to 'word')
        conversion_engine = config.get('conversion_engine', 'word').lower()
        
        # Validate conversion engine choice
        if conversion_engine not in ['word', 'libreoffice']:
            print(f"Warning: Invalid conversion_engine '{conversion_engine}'. Must be 'word' or 'libreoffice'. Defaulting to 'word'.")
            conversion_engine = 'word'
        
        # Check if LibreOffice is requested but not available
        if conversion_engine == 'libreoffice':
            if not libreoffice_module_available:
                print("ERROR: LibreOffice conversion requested but libreoffice_docx_to_pdf.py module not found.")
                print("Please ensure libreoffice_docx_to_pdf.py is in the same directory.")
                if dependencies_ok:
                    print("Falling back to Word conversion.")
                    conversion_engine = 'word'
                else:
                    print("Cannot continue: Neither LibreOffice nor Word conversion is available.")
                    sys.exit(1)
            elif not libreoffice_installed:
                print("ERROR: LibreOffice conversion requested but LibreOffice is not installed or not found.")
                print("Please install LibreOffice and ensure it's in the system PATH.")
                if dependencies_ok:
                    print("Falling back to Word conversion.")
                    conversion_engine = 'word'
                else:
                    print("Cannot continue: Neither LibreOffice nor Word conversion is available.")
                    sys.exit(1)
        
        # If using Word conversion, ensure dependencies are installed
        if conversion_engine == 'word' and not dependencies_ok:
            if libreoffice_available:
                print("\nSwitching to LibreOffice conversion engine due to missing dependencies for Word.")
                conversion_engine = 'libreoffice'
            else:
                print("\nERROR: Word conversion unavailable and LibreOffice fallback failed.")
                if not libreoffice_module_available:
                    print("libreoffice_docx_to_pdf.py module not found.")
                elif not libreoffice_installed:
                    print("LibreOffice is not installed or not found on your system.")
                
                print("\nPlease either:")
                print("1. Install the required packages for Word automation:")
                if platform.system() == 'Windows':
                    print("   pip install pywin32")
                elif platform.system() == 'Darwin':
                    print("   pip install pyobjc")
                print("2. Install LibreOffice and ensure libreoffice_docx_to_pdf.py is available")
                sys.exit(1)
        
        # Get max_threads configuration if specified
        max_workers = 1  # Default to 1 thread (recommended)
        if 'max_threads' in config:
            try:
                max_workers = int(config['max_threads'])
                if max_workers < 1:
                    print(f"Warning: Invalid max_threads value '{max_workers}'. Must be at least 1. Using 1 thread.")
                    max_workers = 1
            except ValueError:
                print(f"Warning: Invalid max_threads value '{config['max_threads']}'. Must be an integer. Using 1 thread.")
        
        # Notify if using more than 1 thread (since single-threaded is usually faster)
        if max_workers > 1:
            print(f"Note: Using {max_workers} threads for conversion (single-threaded mode is usually faster)")
        
        # Check if running on supported OS for Word conversion
        if conversion_engine == 'word' and platform.system() not in ['Windows', 'Darwin']:
            print("Error: Word conversion only supports Windows and macOS")
            print("Consider using LibreOffice conversion for this platform.")
            sys.exit(1)
        
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
        
        # Step 2: Convert Word documents to PDF using selected engine
        print(f"\nStep 2: Converting to PDF using {conversion_engine.capitalize()}...")
        if conversion_engine == 'libreoffice':
            pdf_success, pdf_total = create_pdfs_libreoffice(output_dir, output_dir, max_workers=max_workers)
        else:  # default to Word
            pdf_success, pdf_total = create_pdfs_word(output_dir, output_dir, max_workers=max_workers)
        
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
        print(f"Conversion engine: {conversion_engine.capitalize()}")
        if max_workers is not None:
            print(f"Threads used: {max_workers}")
        if keep_word:
            print(f"Word files directory: {os.path.abspath(output_dir)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

if __name__ == "__main__":
    main() 