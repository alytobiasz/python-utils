"""
Text Search Excerpts Extractor

This script finds excerpts containing multiple search terms in text file(s) and outputs the results to a CSV file.
The script performs case-insensitive searches for each term and creates a new CSV file with a timestamp in
the filename.

IMPORTANT: For performance, this script is designed to perform searches on ALL-LOWERCASE text files only.
Search terms can be provided in any case.

Input:
    - A text file OR directory containing text files to search within
    - A string of forward-slash-separated search terms (e.g., "term1/term2/term three")
      The script will find excerpts containing ANY of these terms

Output:
    - A CSV file named 'search_excerpts_YYYYMMDD_HHMMSS.csv' containing:
        * First column: Name of each input file
        * Second column: All excerpts from that file containing any search term

Excerpt Logic:
    - Natural language: Extract the entire sentence containing the term
    - Non-natural language: Extract the line containing the term plus one line before and after

Usage:
    python text_search_excerpts.py path/to/text_file_or_directory "term1/term2/term three"

Example output CSV:
    Filename,Excerpts
    file1.txt,"Sentence with term1 here. Another sentence with term2 found."
    file2.txt,"Line before\nLine with term three\nLine after"
"""

import argparse
import csv
import os
from datetime import datetime
import time
import re

def is_natural_language_line(line):
    """
    Determine if a line contains natural language text.
    Heuristic: Natural language has spaces, common punctuation, and reasonable word patterns.
    """
    # Remove extra whitespace
    line = line.strip()
    
    # Empty lines are not natural language
    if not line:
        return False
    
    # Count words (sequences of letters/apostrophes)
    words = re.findall(r"[a-zA-Z']+", line)
    
    # If no words, not natural language
    if not words:
        return False
    
    # Natural language indicators:
    # 1. Has multiple words
    # 2. Has sentence-ending punctuation or common punctuation
    # 3. Average word length is reasonable (2-15 characters)
    # 4. Has spaces between words
    
    has_multiple_words = len(words) >= 2
    has_sentence_punctuation = any(punct in line for punct in '.!?,:;')
    has_spaces = ' ' in line
    avg_word_length = sum(len(word) for word in words) / len(words) if words else 0
    reasonable_word_length = 2 <= avg_word_length <= 15
    
    # Consider it natural language if it meets most criteria
    score = sum([has_multiple_words, has_sentence_punctuation, has_spaces, reasonable_word_length])
    return score >= 2

def extract_sentence_containing_term(text, term_match_start, term_match_end):
    """Extract the complete sentence containing a matched term."""
    # Find sentence boundaries - only actual sentence endings, not newlines
    sentence_endings = ['.', '!', '?']
    
    # Find start of sentence (work backwards from match)
    sentence_start = 0
    for i in range(term_match_start - 1, -1, -1):
        if text[i] in sentence_endings:
            sentence_start = i + 1
            break
    
    # Find end of sentence (work forwards from match)
    sentence_end = len(text)
    for i in range(term_match_end, len(text)):
        if text[i] in sentence_endings:
            sentence_end = i + 1
            break
    
    sentence = text[sentence_start:sentence_end].strip()
    # Remove all types of line breaks and normalize whitespace for CSV output
    # This handles \n, \r\n, \r, and multiple spaces/tabs
    sentence = re.sub(r'\s+', ' ', sentence).strip()
    return sentence

def extract_line_context(lines, line_index):
    """Extract the line containing the term plus one line before and after."""
    start_line = max(0, line_index - 1)
    end_line = min(len(lines), line_index + 2)
    
    context_lines = lines[start_line:end_line]
    # Join lines with spaces and normalize all whitespace
    context = ' '.join(line.strip() for line in context_lines if line.strip())
    # Remove all types of line breaks and normalize whitespace
    context = re.sub(r'\s+', ' ', context).strip()
    return context

def find_excerpts(filepath, search_terms):
    """Find all excerpts containing any of the search terms."""
    try:
        with open(filepath, 'r', encoding='utf-8') as file:
            content = file.read()
            lines = content.split('\n')
            
        # Convert to lowercase for case-insensitive search
        content_lower = content.lower()
        
        excerpts = []
        processed_excerpts = set()  # Track actual excerpt text to avoid true duplicates
        
        for term in search_terms:
            # Create regex pattern for whole word matching
            pattern = r'\b' + re.escape(term.lower()) + r'\b'
            
            for match in re.finditer(pattern, content_lower):
                match_start = match.start()
                match_end = match.end()
                
                # Find which line contains this match
                char_count = 0
                line_index = 0
                for i, line in enumerate(lines):
                    if char_count <= match_start < char_count + len(line) + 1:  # +1 for newline
                        line_index = i
                        break
                    char_count += len(line) + 1
                
                # Determine if this line contains natural language
                if line_index < len(lines) and is_natural_language_line(lines[line_index]):
                    # Natural language: extract full sentence
                    excerpt = extract_sentence_containing_term(content, match_start, match_end)
                else:
                    # Non-natural language: extract line context
                    excerpt = extract_line_context(lines, line_index)
                
                # Only add if we haven't seen this exact excerpt before
                excerpt_cleaned = excerpt.strip()
                if excerpt_cleaned and excerpt_cleaned not in processed_excerpts:
                    excerpts.append(excerpt_cleaned)
                    processed_excerpts.add(excerpt_cleaned)
        
        return excerpts
        
    except FileNotFoundError:
        print(f"Error: File not found at '{filepath}'")
        return None
    except Exception as e:
        print(f"Error reading file: {e}")
        return None

def parse_search_terms(terms_string):
    """Parse the forward-slash-separated search terms string."""
    return [term.strip() for term in terms_string.split('/') if term.strip()]

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

def write_results_to_csv(file_results):
    """Write the results to a CSV file."""
    # Create data directory if it doesn't exist
    if not os.path.exists('data'):
        os.makedirs('data')
    
    # Generate output filename with timestamp
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    output_file = os.path.join('data', f'search_excerpts_{timestamp}.csv')
    
    # Find the maximum number of excerpts in any single file
    max_excerpts = max(len(excerpts) for _, excerpts in file_results) if file_results else 0
    
    # Write to CSV
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        
        # Write header - filename plus one column per excerpt
        header = ['Filename'] + [f'Excerpt {i+1}' for i in range(max_excerpts)]
        writer.writerow(header)
        
        # Write data for each file
        for filepath, excerpts in file_results:
            filename = os.path.basename(filepath)
            # Create row with filename and only the excerpts that exist (no padding)
            row = [filename] + excerpts
            writer.writerow(row)
    
    return output_file

def main():
    parser = argparse.ArgumentParser(description='Extract excerpts containing search terms from text file(s)')
    parser.add_argument('filepath', help='Path to text file or directory containing text files')
    parser.add_argument('search_terms', help='Forward-slash-separated search terms (e.g., "term1/term2/term three")')
    args = parser.parse_args()
    
    search_terms = parse_search_terms(args.search_terms)
    if not search_terms:
        print("Error: No valid search terms provided")
        return
    
    print(f"Searching for terms: {', '.join(search_terms)}")
    
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
        
        excerpts = find_excerpts(filepath, search_terms)
        if excerpts is not None:
            elapsed_time = time.time() - start_time
            print(f"Found {len(excerpts)} excerpts in {elapsed_time:.2f} seconds")
            file_results.append((filepath, excerpts))
        else:
            elapsed_time = time.time() - start_time
            print(f"Failed after {elapsed_time:.2f} seconds")
    
    if file_results:
        output_file = write_results_to_csv(file_results)
        total_time = time.time() - total_start_time
        print(f"\nSuccessfully processed {len(file_results)} out of {len(files_to_process)} files")
        print(f"Total processing time: {total_time:.2f} seconds")
        print(f"Results have been written to '{output_file}'")
    else:
        print("No files were successfully processed")

if __name__ == "__main__":
    main() 