#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
DOCX to PDF Converter (LibreOffice Edition)

This script converts all Word (.docx) files in a specified directory to PDF format
using LibreOffice in headless mode (no UI). This is typically faster and more lightweight
than using Microsoft Word for conversions.

Requirements:
    - LibreOffice must be installed
    - The 'soffice' or 'libreoffice' command must be in your PATH
    - No additional Python packages required

Usage:
    python libreoffice_docx_to_pdf.py <directory_path> [max_workers]

Example:
    python libreoffice_docx_to_pdf.py /path/to/documents 4

Note:
    - The script will maintain the original .docx files
    - PDFs will be created in a 'pdf_exports' subdirectory
    - If a PDF with the same name already exists, it will be overwritten
    - Files in subdirectories are not processed (only top-level directory)
    - LibreOffice is started in headless mode for performance
"""

import sys
import os
import time
import platform
import subprocess
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

def is_libreoffice_installed():
    """
    Simple check if LibreOffice is installed and available on the system.
    
    Returns:
        bool: True if LibreOffice is found in PATH, False otherwise
    """
    try:
        # Use appropriate command based on platform
        if platform.system() == 'Windows':
            result = subprocess.run(['where', 'soffice'], capture_output=True, text=True)
        else:  # macOS, Linux, etc.
            result = subprocess.run(['which', 'soffice'], capture_output=True, text=True)
        
        # Return True if command was successful (return code 0)
        return result.returncode == 0
    except Exception:
        # If any error occurs, assume LibreOffice is not available
        return False

def get_libreoffice_cmd():
    """Determine the correct LibreOffice command for the current platform."""
    if platform.system() == 'Windows':
        # Try to find LibreOffice in common installation locations
        potential_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for path in potential_paths:
            if os.path.exists(path):
                return path
        # Fall back to hoping it's in the PATH
        return "soffice.exe"
    elif platform.system() == 'Darwin':  # macOS
        # Check common macOS locations
        potential_paths = [
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/Applications/LibreOffice.app/Contents/MacOS/libreoffice",
        ]
        for path in potential_paths:
            if os.path.exists(path):
                return path
        # Fall back to which command
        try:
            result = subprocess.run(["which", "libreoffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except:
            pass
        # Last resort
        return "libreoffice"
    else:  # Linux and others
        # Try to find in PATH
        try:
            result = subprocess.run(["which", "libreoffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
            result = subprocess.run(["which", "soffice"], capture_output=True, text=True)
            if result.returncode == 0 and result.stdout.strip():
                return result.stdout.strip()
        except:
            pass
        # Default command if we can't find it
        return "libreoffice"

def convert_batch_with_libreoffice(batch_info):
    """
    Convert a batch of DOCX files to PDF using LibreOffice in headless mode.
    
    Args:
        batch_info: Tuple of (batch_number, docx_files, pdf_dir, total_files, libreoffice_cmd)
        
    Returns:
        Tuple of (batch_results, success_count)
    """
    batch_number, file_batch, pdf_dir, total_files, libreoffice_cmd = batch_info
    batch_results = []
    success_count = 0
    
    # Create a unique user profile directory for this batch to avoid conflicts
    user_profile_dir = os.path.join(os.getcwd(), f"libreoffice_profile_{batch_number}_{os.getpid()}")
    os.makedirs(user_profile_dir, exist_ok=True)
    
    # Use a list to collect files for batch processing
    batch_files = []
    for index, docx_path in file_batch:
        batch_files.append((index, docx_path))
    
    try:
        # First convert files in a batch (LibreOffice can handle multiple files at once)
        file_paths = [path for _, path in batch_files]
        
        start_time = time.time()
        
        # Prepare the command with all files
        cmd = [
            libreoffice_cmd,
            "--headless",
            "--convert-to", "pdf",
            "--outdir", pdf_dir,
            "-env:UserInstallation=file://" + user_profile_dir.replace("\\", "/")
        ] + file_paths
        
        print(f"Batch {batch_number}: Starting conversion of {len(batch_files)} files...")
        
        # Run the conversion
        process = subprocess.run(cmd, capture_output=True, text=True)
        
        batch_time = time.time() - start_time
        
        # Check results for each file
        for index, docx_path in batch_files:
            pdf_filename = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            pdf_path = os.path.join(pdf_dir, pdf_filename)
            
            if os.path.exists(pdf_path):
                success_count += 1
                message = f"[{index}/{total_files}] Successfully converted {os.path.basename(docx_path)}"
                batch_results.append((index, True, message))
                print(message)
            else:
                message = f"[{index}/{total_files}] Failed to convert {os.path.basename(docx_path)}"
                batch_results.append((index, False, message))
                print(message)
        
        if process.returncode != 0:
            print(f"Batch {batch_number} error: {process.stderr}")
        
        print(f"Batch {batch_number} completed in {batch_time:.1f} seconds")
        
    except Exception as e:
        print(f"Error processing batch {batch_number}: {str(e)}")
    finally:
        # Clean up the temporary user profile
        try:
            if os.path.exists(user_profile_dir):
                shutil.rmtree(user_profile_dir, ignore_errors=True)
        except:
            pass
    
    return batch_results, success_count

def convert_docx_to_pdf(docx_files, pdf_dir, max_workers=None):
    """
    Convert multiple DOCX files to PDF using LibreOffice in parallel.
    
    Args:
        docx_files: List of DOCX file paths
        pdf_dir: Output directory for PDF files
        max_workers: Number of parallel LibreOffice instances to use
        
    Returns:
        Tuple of (success_count, total_files)
    """
    total_files = len(docx_files)
    success_count = 0
    results = []
    
    # Find LibreOffice command
    libreoffice_cmd = get_libreoffice_cmd()
    print(f"Using LibreOffice command: {libreoffice_cmd}")
    
    # Check if LibreOffice is available
    try:
        version_cmd = [libreoffice_cmd, "--version"]
        result = subprocess.run(version_cmd, capture_output=True, text=True)
        if result.returncode == 0:
            print(f"LibreOffice version: {result.stdout.strip()}")
        else:
            print("Warning: Could not determine LibreOffice version")
    except Exception as e:
        print(f"Warning: Error checking LibreOffice: {str(e)}")
    
    try:
        # Create indexed task list
        indexed_files = [(i+1, docx_path) for i, docx_path in enumerate(docx_files)]
        
        # Split files into batches for each worker thread
        batch_size = max(1, min(50, (total_files + max_workers - 1) // max_workers))  # Max 50 files per batch
        batches = [indexed_files[i:i+batch_size] for i in range(0, len(indexed_files), batch_size)]
        
        print(f"Processing {total_files} files in {len(batches)} batches with {max_workers} workers")
        print(f"Each worker will process approximately {batch_size} files per batch")
        
        # Prepare batch information
        batch_info = [(i+1, batch, pdf_dir, total_files, libreoffice_cmd) 
                     for i, batch in enumerate(batches)]
        
        # Process batches in parallel
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all batches and collect futures
            future_to_batch = {executor.submit(convert_batch_with_libreoffice, info): info 
                             for info in batch_info}
            
            # Process results as they complete
            for future in as_completed(future_to_batch):
                batch_results, batch_success = future.result()
                results.extend(batch_results)
                success_count += batch_success
    
    except Exception as e:
        print(f"Error in conversion process: {str(e)}")
    
    # Make sure to kill any rogue LibreOffice processes
    try:
        if platform.system() == 'Windows':
            subprocess.run(["taskkill", "/F", "/IM", "soffice.exe", "/T"], 
                         capture_output=True)
        else:
            subprocess.run(["pkill", "soffice.bin"], capture_output=True)
            subprocess.run(["pkill", "libreoffice"], capture_output=True)
    except:
        pass
    
    # Sort results by index for consistent output
    results.sort(key=lambda x: x[0])
    
    return success_count, total_files

def create_pdfs(input_dir, pdf_dir=None, max_workers=None):
    """
    Convert all Word (.docx) files in a directory to PDF format using LibreOffice.
    
    Args:
        input_dir (str): Path to the directory containing Word documents
        pdf_dir (str, optional): Directory for PDF output. If None, creates 'pdf_exports' subdirectory
        max_workers (int, optional): Maximum number of parallel LibreOffice instances to use.
                                  If None, uses a default based on CPU count.
        
    Returns:
        tuple: (success_count, total_files) indicating number of successfully converted files
    """
    try:
        # Verify directory exists
        if not os.path.isdir(input_dir):
            raise ValueError(f"Directory not found: {input_dir}")
        
        # Set up PDF output directory
        if pdf_dir is None:
            pdf_dir = os.path.join(input_dir, 'pdf_exports')
        os.makedirs(pdf_dir, exist_ok=True)
        
        # Find all .docx files in the directory
        docx_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) 
                     if f.endswith('.docx') and os.path.isfile(os.path.join(input_dir, f))]
        
        if not docx_files:
            print("No .docx files found in the specified directory.")
            return 0, 0
        
        print(f"\nFound {len(docx_files)} .docx files to process")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
        # Determine number of workers based on system capabilities if not specified
        if max_workers is None:
            import multiprocessing
            # For LibreOffice, using more instances can actually slow things down
            # due to resource contention, so we limit to half the CPU cores or 4
            max_workers = min(max(1, multiprocessing.cpu_count() // 2), 4)
        
        # Track overall start time
        overall_start = time.time()
        
        # Convert files using LibreOffice
        success_count, total_files = convert_docx_to_pdf(docx_files, pdf_dir, max_workers)
        
        # Calculate overall time
        overall_time = time.time() - overall_start
        
        # Print summary
        print("\nProcessing Summary:")
        print(f"Total time: {overall_time:.1f} seconds")
        print(f"Average time per document: {overall_time/len(docx_files):.1f} seconds")
        print(f"Documents per minute: {(len(docx_files) / overall_time) * 60:.1f}")
        print(f"Total files processed: {success_count}/{len(docx_files)}")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
        return success_count, total_files
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Usage: python libreoffice_docx_to_pdf.py <directory_path> [max_workers]")
        sys.exit(1)
    
    # Parse arguments
    input_dir = sys.argv[1]
    max_workers = int(sys.argv[2]) if len(sys.argv) == 3 else None
    
    create_pdfs(input_dir, max_workers=max_workers)

if __name__ == "__main__":
    main() 