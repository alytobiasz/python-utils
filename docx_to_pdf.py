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
    - For optimal performance, the script reuses a single Word instance for all conversions
"""

import sys
import os
import time
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

def convert_to_pdf_windows_batch(docx_files, pdf_dir, max_workers=4):
    """Convert multiple Word documents to PDF using a pool of persistent Word instances."""
    total_files = len(docx_files)
    success_count = 0
    results = []
    
    def process_document_batch(file_batch):
        """Process a batch of documents using a single Word instance."""
        thread_results = []
        thread_success = 0
        
        import win32com.client
        import pywintypes
        
        word = None
        try:
            # Create one Word instance for this entire batch
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False  # Hide Word
            
            # Process each document in this batch with the same Word instance
            for index, docx_path in file_batch:
                start_time = time.time()
                try:
                    doc = word.Documents.Open(docx_path)
                    pdf_path = os.path.join(pdf_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
                    
                    # PDF export options (17 is the PDF format code)
                    doc.SaveAs(pdf_path, FileFormat=17)
                    doc.Close(SaveChanges=False)
                    
                    elapsed_time = time.time() - start_time
                    message = f"[{index}/{total_files}] Successfully converted {os.path.basename(docx_path)} in {elapsed_time:.1f} seconds"
                    thread_success += 1
                    thread_results.append((index, True, message))
                    print(message)  # Print progress in real-time
                except pywintypes.com_error as e:
                    elapsed_time = time.time() - start_time
                    message = f"[{index}/{total_files}] COM Error converting {os.path.basename(docx_path)}: {str(e)} in {elapsed_time:.1f} seconds"
                    thread_results.append((index, False, message))
                    print(message)  # Print progress in real-time
                except Exception as e:
                    elapsed_time = time.time() - start_time
                    message = f"[{index}/{total_files}] Error converting {os.path.basename(docx_path)}: {str(e)} in {elapsed_time:.1f} seconds"
                    thread_results.append((index, False, message))
                    print(message)  # Print progress in real-time
        finally:
            # Clean up Word instance when the entire batch is done
            if word:
                try:
                    word.Quit()
                except:
                    pass
        
        return thread_results, thread_success
    
    try:
        print(f"Starting conversion with {max_workers} persistent Word instances...")
        
        # Create indexed task list
        indexed_files = [(i+1, docx_path) for i, docx_path in enumerate(docx_files)]
        
        # Split files into batches for each worker thread
        batch_size = (len(docx_files) + max_workers - 1) // max_workers  # Ceiling division
        batches = [indexed_files[i:i+batch_size] for i in range(0, len(indexed_files), batch_size)]
        
        # Process batches in parallel with ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all batches and collect futures
            future_to_batch = {executor.submit(process_document_batch, batch): i for i, batch in enumerate(batches)}
            
            # Process results as they complete
            for future in as_completed(future_to_batch):
                batch_results, batch_success = future.result()
                results.extend(batch_results)
                success_count += batch_success
    
    except Exception as e:
        print(f"Error in parallel conversion: {str(e)}")
    
    # Sort results by index for consistent output order
    results.sort(key=lambda x: x[0])
    
    return success_count, total_files

def convert_to_pdf_macos_batch(docx_files, pdf_dir, max_workers=4):
    """Convert multiple Word documents to PDF using a pool of persistent Word instances."""
    total_files = len(docx_files)
    success_count = 0
    results = []
    
    def process_document_batch(file_batch):
        """Process a batch of documents using a single Word instance."""
        thread_results = []
        thread_success = 0
        
        import subprocess
        
        try:
            # Check if Word is already running in this thread
            check_running_script = '''
                tell application "System Events"
                    set isRunning to exists (processes where name is "Microsoft Word")
                end tell
            '''
            
            process = subprocess.run(['osascript', '-e', check_running_script], capture_output=True, text=True)
            was_running = process.stdout.strip() == "true"
            
            # Start Word for this batch
            subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to activate'], capture_output=True)
            time.sleep(2)  # Give Word time to start
            
            # Process all documents in this batch with one Word instance
            for index, docx_path in file_batch:
                start_time = time.time()
                
                # Convert paths to absolute POSIX paths for AppleScript
                abs_docx_path = os.path.abspath(docx_path).replace('\\', '/')
                pdf_path = os.path.join(pdf_dir, os.path.splitext(os.path.basename(docx_path))[0] + ".pdf")
                abs_pdf_path = os.path.abspath(pdf_path).replace('\\', '/')
                
                # Process a single document without quitting Word
                script = f'''
                    tell application "Microsoft Word"
                        try
                            set docPath to POSIX file "{abs_docx_path}" as alias
                            set pdfPath to POSIX file "{abs_pdf_path}" as string
                            open docPath
                            set docName to name of active document
                            save as active document file format format PDF file name pdfPath
                            
                            # Safely close the document
                            try
                                close active document saving no
                            on error closeErr
                                try
                                    set docList to documents whose name is docName
                                    if (count of docList) > 0 then
                                        close document docName saving no
                                    end if
                                on error
                                    # Just continue, we'll clean up at the end
                                end try
                            end try
                            
                            return "success"
                        on error errMsg
                            try
                                close active document saving no
                            end try
                            return "error: " & errMsg
                        end try
                    end tell
                '''
                
                process = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
                elapsed_time = time.time() - start_time
                
                # Check results
                if os.path.exists(pdf_path):
                    message = f"[{index}/{total_files}] Successfully converted {os.path.basename(docx_path)} in {elapsed_time:.1f} seconds"
                    thread_success += 1
                    thread_results.append((index, True, message))
                    print(message)  # Print progress in real-time
                else:
                    error_msg = process.stdout.strip()
                    if not error_msg.startswith("error:"):
                        error_msg = f"error: PDF not created - {error_msg}"
                    message = f"[{index}/{total_files}] Failed to convert {os.path.basename(docx_path)}: {error_msg} in {elapsed_time:.1f} seconds"
                    thread_results.append((index, False, message))
                    print(message)  # Print progress in real-time
                
                # Try to clean up any open documents periodically to prevent accumulation
                if index % 5 == 0:
                    try:
                        cleanup_script = '''
                            tell application "Microsoft Word"
                                if (count of documents) > 0 then
                                    close all saving no
                                end if
                            end tell
                        '''
                        subprocess.run(['osascript', '-e', cleanup_script], capture_output=True)
                    except:
                        pass
            
            # Quit Word only if it wasn't running before we started
            if not was_running:
                subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'], capture_output=True)
        
        except Exception as e:
            message = f"Error in batch processing: {str(e)}"
            print(message)
            thread_results.append((0, False, message))
            
            # Make sure Word is closed if there was an error
            try:
                subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'], capture_output=True)
            except:
                pass
        
        return thread_results, thread_success
    
    try:
        print(f"Starting conversion with {max_workers} persistent Word instances...")
        
        # Create indexed task list
        indexed_files = [(i+1, docx_path) for i, docx_path in enumerate(docx_files)]
        
        # Split files into batches for each worker thread
        batch_size = (len(docx_files) + max_workers - 1) // max_workers  # Ceiling division
        batches = [indexed_files[i:i+batch_size] for i in range(0, len(indexed_files), batch_size)]
        
        # Process batches in parallel with ThreadPoolExecutor
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all batches and collect futures
            future_to_batch = {executor.submit(process_document_batch, batch): i for i, batch in enumerate(batches)}
            
            # Process results as they complete
            for future in as_completed(future_to_batch):
                batch_results, batch_success = future.result()
                results.extend(batch_results)
                success_count += batch_success
    
    except Exception as e:
        print(f"Error in parallel conversion: {str(e)}")
    
    # Make sure all Word instances are closed
    try:
        import subprocess
        subprocess.run(['osascript', '-e', 'tell application "Microsoft Word" to quit'], 
                       capture_output=True, text=True)
    except:
        pass
    
    # Sort results by index for consistent output order
    results.sort(key=lambda x: x[0])
    
    return success_count, total_files

# Keep original functions for backward compatibility
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

def create_pdfs(input_dir, pdf_dir=None, max_workers=None):
    """
    Convert all Word (.docx) files in a directory to PDF format.
    Uses multithreaded processing with persistent Word instances for optimal performance.
    
    Args:
        input_dir (str): Path to the directory containing Word documents
        pdf_dir (str, optional): Directory for PDF output. If None, creates 'pdf_exports' subdirectory
        max_workers (int, optional): Maximum number of parallel Word instances to use. 
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
            # For persistent instances, limit to available CPU cores or 4, whichever is smaller
            max_workers = min(multiprocessing.cpu_count(), 4)
        
        # Track overall start time
        overall_start = time.time()
        
        # Using optimized multithreaded conversion with persistent Word instances
        if platform.system() == 'Windows':
            print(f"\nUsing {max_workers} persistent Word instances, each handling multiple documents...")
            success_count, total_files = convert_to_pdf_windows_batch(docx_files, pdf_dir, max_workers)
        elif platform.system() == 'Darwin':  # macOS
            print(f"\nUsing {max_workers} persistent Word instances, each handling multiple documents...")
            success_count, total_files = convert_to_pdf_macos_batch(docx_files, pdf_dir, max_workers)
        else:
            raise ValueError("Unsupported operating system")
        
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
        print("Usage: python docx_to_pdf.py <directory_path> [max_workers]")
        sys.exit(1)
    
    # Check if running on supported OS
    if platform.system() not in ['Windows', 'Darwin']:
        print("Error: This script only supports Windows and macOS")
        sys.exit(1)
    
    # Parse arguments
    input_dir = sys.argv[1]
    max_workers = int(sys.argv[2]) if len(sys.argv) == 3 else None
    
    create_pdfs(input_dir, max_workers=max_workers)

if __name__ == "__main__":
    main() 