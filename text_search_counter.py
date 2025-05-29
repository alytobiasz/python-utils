"""
Text Search Counter

This script counts occurrences of multiple search term groups in text file(s) and outputs the results to a CSV file.
The script performs case-insensitive searches for each term and creates a new CSV file with a timestamp in
the filename.

IMPORTANT: For performance, this script is designed to perform searches on ALL-LOWERCASE text files only.
Search terms can be provided in any case.

Input:
    - A text file OR directory containing text files to search within
    - A file containing forward-slash-separated groups of search terms (one group per line)
      Each line represents an OR condition - it counts occurrences of ANY term in the group

Output:
    - A CSV file named 'search_results_YYYYMMDD_HHMMSS.csv' containing:
        * First column: Name of each input file
        * Additional columns: Count for each search term group in that file

Usage:
    python text_search.py path/to/text_file_or_directory path/to/search_terms.txt

Example search_terms.txt:
    hello/hi/hey
    world/earth/globe
    python/programming/code

Example output CSV:
    Filename,hello/hi/hey,world/earth/globe,python/programming/code
    file1.txt,8,3,15
    file2.txt,2,0,7
    file3.txt,5,1,12
"""

import argparse
import csv
import os
from datetime import datetime
import time
import re

def count_occurrences(filepath, term_groups):
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            # Convert content to lowercase for case-insensitive search
            content = file.read().lower()
            
            results = {}
            for group_name, terms in term_groups:
                # Count occurrences of any term in the group (case-insensitive, whole words only)
                total_count = 0
                for term in terms:
                    # Create regex pattern for whole word matching
                    pattern = r'\b' + re.escape(term.lower()) + r'\b[.,!?:;"\'\)\]\}]*'
                    total_count += len(re.findall(pattern, content))
                results[group_name] = total_count
                
            return results
            
    except FileNotFoundError:
        print(f"Error: File not found at '{filepath}'")
        return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

def read_search_terms(terms_filepath):
    try:
        with open(terms_filepath, 'r', encoding='utf-8') as file:
            # Read non-empty lines and split into groups
            term_groups = []
            for line in file:
                if line.strip():
                    terms = [term.strip() for term in line.strip().split('/')]
                    group_name = line.strip()
                    term_groups.append((group_name, terms))
            return term_groups
    except FileNotFoundError:
        print(f"Error: Search terms file not found at '{terms_filepath}'")
        return None
    except Exception as e:
        print(f"Error reading search terms file: {e}")
        return None

def get_files_to_process(path):
    """Get list of files to process from path (single file or directory)"""
    if os.path.isfile(path):
        return [path]
    elif os.path.isdir(path):
        # Get all files in directory
        files = []
        for filename in os.listdir(path):
            filepath = os.path.join(path, filename)
            if os.path.isfile(filepath):
                files.append(filepath)
        return sorted(files)  # Sort for consistent ordering
    return []

def write_results_to_csv(file_results, term_groups):
    # Create data directory if it doesn't exist
    if not os.path.exists('data'):
        os.makedirs('data')
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join('data', f'search_results_{timestamp}.csv')
    
    # Write to CSV, creating a new file
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        # Write header - use group names
        header = ['Filename'] + [group_name for group_name, _ in term_groups]
        writer.writerow(header)
        
        # Write data for each file
        for filepath, results in file_results:
            filename = os.path.basename(filepath)
            row = [filename] + [results[group_name] for group_name, _ in term_groups]
            writer.writerow(row)
    
    return output_file

def main():
    parser = argparse.ArgumentParser(description='Count occurrences of search term groups in text file(s)')
    parser.add_argument('filepath', help='Path to text file or directory containing text files')
    parser.add_argument('terms_file', help='Path to file containing forward-slash-separated search term groups (one group per line)')
    args = parser.parse_args()
    
    term_groups = read_search_terms(args.terms_file)
    if not term_groups:
        return
    
    files_to_process = get_files_to_process(args.filepath)
    if not files_to_process:
        print(f"Error: No files found at '{args.filepath}'")
        return
    
    print(f"\nProcessing {len(files_to_process)} files...")
    total_start_time = time.time()
    
    # Process each file and collect results
    file_results = []
    for i, filepath in enumerate(files_to_process, 1):
        print(f"\nProcessing file {i}/{len(files_to_process)}: {filepath}")
        start_time = time.time()
        
        results = count_occurrences(filepath, term_groups)
        if results:
            elapsed_time = time.time() - start_time
            print(f"Completed in {elapsed_time:.2f} seconds")
            file_results.append((filepath, results))
        else:
            elapsed_time = time.time() - start_time
            print(f"Failed after {elapsed_time:.2f} seconds")
    
    if file_results:
        output_file = write_results_to_csv(file_results, term_groups)
        total_time = time.time() - total_start_time
        print(f"\nSuccessfully processed {len(file_results)} out of {len(files_to_process)} files")
        print(f"Total processing time: {total_time:.2f} seconds")
        print(f"Results have been written to '{output_file}'")

if __name__ == "__main__":
    main() 