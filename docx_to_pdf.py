#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX to PDF Converter

This script converts all Word (.docx) files in a specified directory to PDF format
using Microsoft Word's native interfaces (no macros required).

Requirements:
    - Microsoft Word must be installed
    - Python packages:
        Windows: pip install pywin32
        macOS: pip install pyobjc

Usage:
    python docx_to_pdf.py <directory_path>

Example:
    python docx_to_pdf.py /path/to/documents

Note:
    - The script will maintain the original .docx files
    - PDFs will be created in a 'pdf_exports' subdirectory
    - If a PDF with the same name already exists, it will be overwritten
    - Files in subdirectories are not processed (only top-level directory)
"""

import sys
import os
import time
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

def convert_to_pdf_windows(docx_path, pdf_dir):
    """Convert Word document to PDF using Windows COM interface."""
    import win32com.client
    import pywintypes
    
    try:
        word = win32com.client.Dispatch("Word.Application")
        doc = word.Documents.Open(docx_path)
        pdf_path = os.path.join(pdf_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        
        # PDF export options (17 is the PDF format code)
        doc.SaveAs(pdf_path, FileFormat=17)
        doc.Close()
        word.Quit()
        
        return True, f"Successfully converted {os.path.basename(docx_path)}"
    except pywintypes.com_error as e:
        return False, f"COM Error converting {os.path.basename(docx_path)}: {str(e)}"
    except Exception as e:
        return False, f"Error converting {os.path.basename(docx_path)}: {str(e)}"
    finally:
        try:
            doc.Close(False)
            word.Quit()
        except:
            pass

def convert_to_pdf_macos(docx_path, pdf_dir):
    """Convert Word document to PDF using AppleScript."""
    try:
        # Convert paths to absolute POSIX paths for AppleScript
        docx_path = os.path.abspath(docx_path).replace('\\', '/')
        pdf_path = os.path.join(pdf_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
        pdf_path = os.path.abspath(pdf_path).replace('\\', '/')
        
        # AppleScript to convert document without using macros
        script = f'''
            tell application "Microsoft Word"
                set docPath to POSIX file "{docx_path}" as alias
                set pdfPath to POSIX file "{pdf_path}" as string
                open docPath
                set theDoc to active document
                save as theDoc file format format PDF file name pdfPath
                close theDoc saving no
                quit
            end tell
        '''
        
        import subprocess
        process = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        
        if process.returncode != 0:
            return False, f"Error converting {os.path.basename(docx_path)}: {process.stderr}"
        
        if not os.path.exists(pdf_path):
            return False, f"PDF file was not created for {os.path.basename(docx_path)}"
            
        return True, f"Successfully converted {os.path.basename(docx_path)}"
    except Exception as e:
        return False, f"Error converting {os.path.basename(docx_path)}: {str(e)}"

def convert_to_pdf(docx_path, pdf_dir):
    """
    Convert a single Word document to PDF using platform-specific method.
    
    Args:
        docx_path (str): Path to the Word document
        pdf_dir (str): Directory where PDF should be saved
        
    Returns:
        tuple: (success, message) where success is a boolean and message is a string
    """
    start_time = time.time()
    
    if platform.system() == 'Windows':
        success, message = convert_to_pdf_windows(docx_path, pdf_dir)
    elif platform.system() == 'Darwin':  # macOS
        success, message = convert_to_pdf_macos(docx_path, pdf_dir)
    else:
        return False, "Unsupported operating system"
    
    elapsed_time = time.time() - start_time
    return success, f"{message} in {elapsed_time:.1f} seconds"

def process_directory(directory):
    """
    Process all .docx files in the specified directory.
    
    Args:
        directory (str): Path to the directory containing Word documents
    """
    try:
        # Verify directory exists
        if not os.path.isdir(directory):
            print(f"Error: Directory not found: {directory}")
            sys.exit(1)
        
        # Create pdf_exports directory
        pdf_dir = os.path.join(directory, 'pdf_exports')
        os.makedirs(pdf_dir, exist_ok=True)
        
        # Find all .docx files in the directory
        docx_files = [os.path.join(directory, f) for f in os.listdir(directory) 
                     if f.endswith('.docx') and os.path.isfile(os.path.join(directory, f))]
        
        if not docx_files:
            print("No .docx files found in the specified directory.")
            return
        
        print(f"\nFound {len(docx_files)} .docx files to process")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
        # Process files sequentially (Word doesn't handle parallel processing well)
        total_start_time = time.time()
        success_count = 0
        
        for i, docx_file in enumerate(docx_files, 1):
            success, message = convert_to_pdf(docx_file, pdf_dir)
            if success:
                success_count += 1
            print(f"[{i}/{len(docx_files)}] {message}")
        
        # Print summary
        total_time = time.time() - total_start_time
        print("\nProcessing Summary:")
        print(f"Total files processed: {success_count}/{len(docx_files)}")
        print(f"Total processing time: {total_time:.1f} seconds")
        print(f"Average time per file: {(total_time/len(docx_files)):.1f} seconds")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

def main():
    """Main function to handle command line arguments and start processing."""
    if len(sys.argv) != 2:
        print("Usage: python docx_to_pdf.py <directory_path>")
        sys.exit(1)
    
    # Check if running on supported OS
    if platform.system() not in ['Windows', 'Darwin']:
        print("Error: This script only supports Windows and macOS")
        sys.exit(1)
    
    directory = sys.argv[1]
    process_directory(directory)

if __name__ == "__main__":
    main() 