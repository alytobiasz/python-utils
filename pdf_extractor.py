"""
PDF Text Extractor

This script extracts text content from PDF files.
Can accept either:
1. A text file containing PDF file paths (one per line), or
2. A directory path containing PDF files

Usage:
    python pdf_extractor.py path/to/your/list.txt
    python pdf_extractor.py path/to/pdf/directory/

Output:
    Creates a directory named 'pdf_extracts_YYYYMMDD_HHMMSS' containing all extracted text files

Requirements:
    pip install PyMuPDF
"""

import argparse
import time
from datetime import datetime
import os
import glob
import fitz  # PyMuPDF

def extract_text_from_pdf_file(pdf_path):
    try:
        with fitz.open(pdf_path) as pdf:
            text = ""
            for page in pdf:
                text += page.get_text() + "\n\n"
            return text.strip()
    
    except FileNotFoundError:
        print(f"Error: File not found at '{pdf_path}'")
        return None
    except Exception as e:
        print(f"Error processing PDF: {e}")
        return None

def get_pdf_paths_from_directory(directory_path):
    """Get all PDF file paths from a directory."""
    pdf_extensions = ['*.pdf', '*.PDF']
    pdf_paths = []
    
    for extension in pdf_extensions:
        pattern = os.path.join(directory_path, extension)
        pdf_paths.extend(glob.glob(pattern))
    
    # Sort paths for consistent processing order
    return sorted(pdf_paths)

def get_pdf_paths_from_file(file_path):
    """Get PDF file paths from a text file (one path per line)."""
    with open(file_path, 'r') as file:
        return [line.strip() for line in file if line.strip()]

def process_pdfs(input_path):
    """Process PDFs from either a directory or a file list."""
    try:
        # Determine if input is a directory or file
        if os.path.isdir(input_path):
            print(f"Processing directory: {input_path}")
            pdf_paths = get_pdf_paths_from_directory(input_path)
            if not pdf_paths:
                print(f"No PDF files found in directory: {input_path}")
                return
        elif os.path.isfile(input_path):
            print(f"Processing file list: {input_path}")
            pdf_paths = get_pdf_paths_from_file(input_path)
        else:
            print(f"Error: '{input_path}' is neither a valid file nor directory")
            return
        
        # Create data directory if it doesn't exist
        if not os.path.exists('data'):
            os.makedirs('data')
        
        # Create timestamped output directory inside data folder
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_dir = os.path.join('data', f'pdf_extracts_{timestamp}')
        os.makedirs(output_dir, exist_ok=True)
        
        success_count = 0
        print(f"\nProcessing {len(pdf_paths)} files...")
        print(f"Output directory: {output_dir}")
        
        total_start_time = time.time()
        
        for i, pdf_path in enumerate(pdf_paths, 1):
            start_time = time.time()
            print(f"\nProcessing file {i}/{len(pdf_paths)}: {pdf_path}")
            
            extracted_text = extract_text_from_pdf_file(pdf_path)
            
            if extracted_text:
                # Generate output filename from last 5 parts of the path
                path_parts = pdf_path.replace('\\', '/').split('/')
                last_parts = path_parts[-5:] if len(path_parts) >= 5 else path_parts
                base_name = '-'.join(last_parts).rsplit('.', 1)[0]
                output_filename = os.path.join(output_dir, f"{base_name}.txt")
                
                # Save to file
                with open(output_filename, "w", encoding="utf-8") as f:
                    f.write(extracted_text)
                    
                elapsed_time = time.time() - start_time
                print(f"Text saved to '{output_filename}'\nCompleted in {elapsed_time:.2f} seconds.")
                success_count += 1
            else:
                elapsed_time = time.time() - start_time
                print(f"Failed after {elapsed_time:.2f} seconds: {pdf_path}")
        
        total_time = time.time() - total_start_time
        print(f"\nProcessing complete. Successfully processed {success_count} out of {len(pdf_paths)} files.")
        print(f"Total processing time: {total_time:.2f} seconds")
        print(f"All output files are in directory: {output_dir}")
        
    except FileNotFoundError:
        print(f"Error: File not found at '{input_path}'")
    except Exception as e:
        print(f"Error processing input: {e}")

def main():
    # Set up command line argument parsing
    parser = argparse.ArgumentParser(description='Extract text from PDF files from a directory or file list')
    parser.add_argument('input_path', help='Path to a directory containing PDFs or a text file containing PDF paths (one per line)')
    args = parser.parse_args()
    
    process_pdfs(args.input_path)

if __name__ == "__main__":
    main() 