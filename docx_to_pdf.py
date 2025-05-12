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
    python docx_to_pdf.py <directory_path> [max_threads]

Example:
    python docx_to_pdf.py /path/to/documents
    python docx_to_pdf.py /path/to/documents 4  # Use 4 threads for conversion

Note:
    - The script will maintain the original .docx files
    - PDFs will be created in a 'pdf_exports' subdirectory
    - If a PDF with the same name already exists, it will be overwritten
    - Files in subdirectories are not processed (only top-level directory)
    - By default, a single Word instance is used (recommended for best performance)
    - Multiple threads can be specified but may not improve performance
"""

import sys
import os
import time
import platform
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

def check_dependencies():
    """Check that platform-specific dependencies are installed."""
    system = platform.system()
    
    if system == 'Windows':
        try:
            # Check if pywin32 is installed
            import win32com.client
            import pythoncom
            return True
        except ImportError:
            print("\nERROR: Required package 'pywin32' is not installed!")
            print("Please install it using the following command:")
            print("    pip install pywin32")
            print("\nThis package is required for Word automation on Windows.")
            return False
    
    elif system == 'Darwin':  # macOS
        try:
            # Check if pyobjc is installed
            import objc
            return True
        except ImportError:
            print("\nERROR: Required package 'pyobjc' is not installed!")
            print("Please install it using the following command:")
            print("    pip install pyobjc")
            print("\nThis package is required for Word automation on macOS.")
            return False
    
    else:
        print(f"\nERROR: Unsupported operating system: {system}")
        print("This script only supports Windows and macOS.")
        if system == 'Linux':
            print("For Linux, consider using the LibreOffice conversion engine instead.")
        return False

# Check dependencies before proceeding
dependencies_ok = check_dependencies()

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
        import pythoncom  # Import pythoncom module for COM initialization
        
        # Initialize COM for this thread
        pythoncom.CoInitialize()
        
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
            # Uninitialize COM for this thread
            pythoncom.CoUninitialize()
        
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
                try:
                    batch_results, batch_success = future.result()
                    results.extend(batch_results)
                    success_count += batch_success
                except Exception as e:
                    print(f"Error in parallel conversion: {str(e)}")
                    # Continue processing other batches even if one fails
    
    except Exception as e:
        print(f"Error in parallel conversion: {str(e)}")
    
    # Sort results by index for consistent output order
    results.sort(key=lambda x: x[0])
    
    # Verify success count against actual results
    verified_success_count = sum(1 for r in results if r[1])
    if verified_success_count != success_count:
        print(f"Warning: Success count mismatch. Calculated: {verified_success_count}, Reported: {success_count}")
        success_count = verified_success_count
    
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

def convert_to_pdf_windows(docx_path, pdf_dir):
    """Convert Word document to PDF using Windows COM interface."""
    import win32com.client
    import pywintypes
    import pythoncom  # Import pythoncom module for COM initialization
    
    # Initialize COM for this thread
    pythoncom.CoInitialize()
    
    word = None
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
            if 'doc' in locals() and doc:
                doc.Close(False)
            if word:
                word.Quit()
        except:
            pass
        
        # Uninitialize COM for this thread
        pythoncom.CoUninitialize()

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

def create_pdfs(input_dir, pdf_dir=None, max_workers=1):
    """
    Convert all Word (.docx) files in a directory to PDF format.
    
    Args:
        input_dir (str): Path to the directory containing Word documents
        pdf_dir (str, optional): Directory for PDF output. If None, creates 'pdf_exports' subdirectory
        max_workers (int, optional): Maximum number of parallel Word instances to use.
                                  Default is 1, which is recommended for optimal performance.
        
    Returns:
        tuple: (success_count, total_files) indicating number of successfully converted files
    """
    success_count = 0
    total_files = 0
    
    try:
        # Verify directory exists
        input_dir = os.path.abspath(input_dir)
        if not os.path.isdir(input_dir):
            raise ValueError(f"Directory not found: {input_dir}")
        
        # Set up PDF output directory
        if pdf_dir is None:
            pdf_dir = os.path.join(input_dir, 'pdf_exports')
        os.makedirs(pdf_dir, exist_ok=True)
        
        # Find all .docx files in the directory
        docx_files = [os.path.join(input_dir, f) for f in os.listdir(input_dir) 
                     if f.endswith('.docx') and os.path.isfile(os.path.join(input_dir, f))]
        
        total_files = len(docx_files)
        if not docx_files:
            print("No .docx files found in the specified directory.")
            return 0, 0
        
        print(f"\nFound {total_files} .docx files to process")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
        # Ensure max_workers is at least 1
        if max_workers is None or max_workers < 1:
            max_workers = 1
            
        # Display appropriate message based on thread count
        if max_workers == 1:
            worker_message = "a single Word instance for all documents"
        else:
            worker_message = f"{max_workers} persistent Word instances, each handling multiple documents"
        
        # Track overall start time
        overall_start = time.time()
        
        # Use the specified number of threads (default is 1)
        try:
            if platform.system() == 'Windows':
                print(f"\nUsing {worker_message}...")
                success_count, _ = convert_to_pdf_windows_batch(docx_files, pdf_dir, max_workers)
            elif platform.system() == 'Darwin':  # macOS
                print(f"\nUsing {worker_message}...")
                success_count, _ = convert_to_pdf_macos_batch(docx_files, pdf_dir, max_workers)
            else:
                raise ValueError("Unsupported operating system")
        except Exception as e:
            print(f"Error in batch conversion: {str(e)}")
            print("Attempting to convert files individually as fallback...")
            
            # Fallback to individual conversion if batch fails
            success_count = 0
            for i, docx_path in enumerate(docx_files):
                try:
                    success, message = convert_to_pdf(docx_path, pdf_dir)
                    if success:
                        success_count += 1
                    print(f"[{i+1}/{total_files}] {message}")
                except Exception as e:
                    print(f"[{i+1}/{total_files}] Error: {str(e)}")
        
        # Verify success count by counting actual PDF files
        created_pdfs = []
        for docx_path in docx_files:
            pdf_name = os.path.splitext(os.path.basename(docx_path))[0] + ".pdf"
            pdf_path = os.path.join(pdf_dir, pdf_name)
            if os.path.exists(pdf_path):
                created_pdfs.append(pdf_path)
        
        actual_pdf_count = len(created_pdfs)
        if actual_pdf_count != success_count:
            print(f"Warning: Success count ({success_count}) doesn't match actual PDF files found ({actual_pdf_count})")
            print("Using actual count of PDF files created for reporting.")
            success_count = actual_pdf_count
        
        # Calculate overall time
        overall_time = time.time() - overall_start
        
        # Print summary
        print("\nProcessing Summary:")
        print(f"Total time: {overall_time:.1f} seconds")
        if total_files > 0:
            print(f"Average time per document: {overall_time/total_files:.1f} seconds")
        print(f"PDF files created: {success_count}/{total_files}")
        print(f"Output directory: {os.path.abspath(pdf_dir)}")
        
        return success_count, total_files
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        return success_count, total_files  # Return what we have so far rather than exiting

def main():
    """Main function to handle command line arguments."""
    if len(sys.argv) < 2 or len(sys.argv) > 3:
        print("Usage: python docx_to_pdf.py <directory_path> [max_threads]")
        sys.exit(1)
    
    # Check if dependencies are installed
    if not dependencies_ok:
        print("\nCannot continue: Required dependencies are missing.")
        sys.exit(1)
    
    # Check if running on supported OS
    if platform.system() not in ['Windows', 'Darwin']:
        print("Error: This script only supports Windows and macOS")
        sys.exit(1)
    
    # Parse arguments
    input_dir = sys.argv[1]
    max_threads = 1  # Default to 1 thread
    
    if len(sys.argv) == 3:
        try:
            max_threads = int(sys.argv[2])
            if max_threads < 1:
                print(f"Warning: Invalid max_threads value '{max_threads}'. Must be at least 1. Using 1 thread.")
                max_threads = 1
        except ValueError:
            print(f"Warning: Invalid max_threads value '{sys.argv[2]}'. Must be an integer. Using 1 thread.")
    
    if max_threads == 1:
        print("Using single-threaded mode (recommended for optimal performance)")
    else:
        print(f"Using {max_threads} threads for conversion (single-threaded mode is usually faster)")
    
    create_pdfs(input_dir, max_workers=max_threads)

if __name__ == "__main__":
    main() 