"""
File Copier

This script copies files listed in a text file to a specified output directory.
Each line in the input file should contain a full path to a file, or a relative 
path to the current working directory.

Usage:
    python file_copier.py <file_list.txt> <output_directory>

Example:
    python file_copier.py files_to_copy.txt output_folder

Note:
    - The script will create the output directory if it doesn't exist
    - Files with the same name will be renamed with a numeric suffix
    - The script maintains the original file names but not the directory structure
"""

import sys
import os
import shutil
import time
from datetime import datetime

def copy_files(file_list_path, output_dir):
    """
    Copy files from the list to the output directory.
    
    Args:
        file_list_path (str): Path to the text file containing file paths
        output_dir (str): Path to the output directory
    """
    try:
        # Create output directory if it doesn't exist
        os.makedirs(output_dir, exist_ok=True)
        
        # Read the file list
        with open(file_list_path, 'r') as f:
            files = [line.strip() for line in f if line.strip()]
        
        if not files:
            print("No files found in the list.")
            return
        
        print(f"\nFound {len(files)} files to copy")
        print(f"Output directory: {os.path.abspath(output_dir)}")
        
        # Process each file
        total_start_time = time.time()
        success_count = 0
        
        for i, source_path in enumerate(files, 1):
            start_time = time.time()
            
            try:
                if not os.path.exists(source_path):
                    print(f"File not found: {source_path}")
                    continue
                
                # Get the base filename without path
                filename = os.path.basename(source_path)
                
                # Create destination path
                dest_path = os.path.join(output_dir, filename)
                
                # Handle duplicate filenames
                counter = 1
                base, ext = os.path.splitext(dest_path)
                while os.path.exists(dest_path):
                    dest_path = f"{base}_{counter}{ext}"
                    counter += 1
                
                # Copy the file
                shutil.copy2(source_path, dest_path)
                
                elapsed_time = time.time() - start_time
                success_count += 1
                print(f"Copied {i}/{len(files)} files: {filename} in {elapsed_time:.1f} seconds")
                
            except Exception as e:
                print(f"Error copying {source_path}: {str(e)}")
        
        # Print summary
        total_time = time.time() - total_start_time
        print("\nProcessing Summary:")
        print(f"Total files copied: {success_count}/{len(files)}")
        print(f"Total processing time: {total_time:.1f} seconds")
        print(f"Average time per file: {(total_time/len(files)):.1f} seconds")
        print(f"Output directory: {os.path.abspath(output_dir)}")
        
    except Exception as e:
        print(f"Error: {str(e)}")
        import traceback
        traceback.print_exc()
        sys.exit(1)

def main():
    if len(sys.argv) != 3:
        print("Usage: python file_copier.py <file_list.txt> <output_directory>")
        sys.exit(1)
    
    file_list_path = sys.argv[1]
    output_dir = sys.argv[2]
    
    copy_files(file_list_path, output_dir)

if __name__ == "__main__":
    main() 