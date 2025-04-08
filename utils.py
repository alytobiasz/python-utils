#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Utility functions for document processing scripts.

This module contains shared utility functions used by multiple document processing scripts:
- docx_template_filler.py
- pdf_form_filler.py
- pdf_template_filler.py

Functions include:
- Common date formatting for consistent display of dates
- Filename sanitization to ensure valid filenames
- Configuration file reading with standardized error handling 
- Unique filename generation to avoid overwrites

IMPORTANT: All scripts now require exact field name matches between Excel headers and PDF fields.
No automatic normalization of field names is performed - field names are case-sensitive and 
space-sensitive.
"""

import datetime as dt
import os
import re

def format_date(value, include_time=True):
    """
    Format a date/datetime value consistently as "Month Day, Year" (e.g., "January 1, 2025").
    If the value has a time component and include_time is True, it will be included.
    
    Args:
        value: A date or datetime object
        include_time: Whether to include time component if present (default: True)
        
    Returns:
        str: The formatted date string
    """
    if value is None:
        return ''
        
    if isinstance(value, (dt.datetime, dt.date)):
        # Special case: If time component exists and needs to be displayed
        has_time = False
        if isinstance(value, dt.datetime):
            has_time = not (value.hour == 0 and value.minute == 0 and value.second == 0)
            
        if has_time and include_time:
            # Default to "Month Day, Year" format with time
            time_format = '%H:%M:%S' if value.second != 0 else '%H:%M'
            month_name = value.strftime('%B')  # Full month name
            day = value.day
            year = value.year
            return f"{month_name} {day}, {year} {value.strftime(time_format)}"
        
        # Default to "Month Day, Year" format
        month_name = value.strftime('%B')  # Full month name
        day = value.day
        year = value.year
        return f"{month_name} {day}, {year}"
    
    # For non-date values, return as string
    return str(value)

def format_excel_cell_date(cell):
    """
    Gets the formatted date value from an Excel cell.
    For dates, standardizes to "Month Day, Year" format (e.g., "January 1, 2025")
    regardless of how they appear in Excel.
    
    Args:
        cell: An openpyxl cell object
        
    Returns:
        str: The formatted value using standardized formatting for dates
    """
    if cell.value is None:
        return ''

    # Handle date values
    if hasattr(cell, 'value') and isinstance(cell.value, (dt.datetime, dt.date)):
        # Check if format has time markers
        number_format = cell.number_format.lower()
        include_time = 'h' in number_format or ':' in number_format
        
        return format_date(cell.value, include_time)
    
    # For non-date values, return as string
    return str(cell.value) 

def sanitize_filename(filename, default_name="document", max_length=200):
    """
    Sanitize a filename by removing invalid characters and ensuring it meets file system requirements.
    
    Args:
        filename (str): The filename to sanitize
        default_name (str): Default name to use if filename is empty after sanitization
        max_length (int): Maximum length for the filename (Windows limits to 255)
        
    Returns:
        str: The sanitized filename
    """
    if not filename:
        return default_name
        
    # Replace invalid characters with underscores
    invalid_chars = r'[<>:"/\\|?*]'
    sanitized = re.sub(invalid_chars, '_', str(filename))
    
    # Remove leading/trailing spaces and dots
    sanitized = sanitized.strip(". ")
    
    # Collapse multiple spaces and underscores to a single underscore
    sanitized = re.sub(r'[\s_]+', '_', sanitized)
    
    # Default filename if empty after sanitization
    if not sanitized:
        sanitized = default_name
    
    # Limit length
    if max_length > 0 and len(sanitized) > max_length:
        sanitized = sanitized[:max_length]
    
    return sanitized

def read_config(config_path, required_fields=None):
    """
    Read configuration from a text file with key=value format.
    
    Args:
        config_path (str): Path to the configuration file
        required_fields (list): List of field names that must be present in the config
        
    Returns:
        dict: Configuration parameters
        
    Raises:
        ValueError: If required fields are missing or file cannot be read
    """
    if required_fields is None:
        required_fields = ['excel_file', 'template', 'output_directory']
        
    config = {}
    
    try:
        with open(config_path, 'r') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    key, value = map(str.strip, line.split('=', 1))
                    config[key] = value
    except Exception as e:
        raise ValueError(f"Error reading config file: {str(e)}")
    
    # Check for required fields
    if required_fields:
        missing = [field for field in required_fields if field not in config]
        if missing:
            raise ValueError(f"Missing required fields in config file: {', '.join(missing)}")
    
    return config

def get_unique_filename(base_path, extension="pdf"):
    """
    Ensure a filename is unique by appending a counter if needed.
    
    Args:
        base_path (str): Base filepath without extension
        extension (str): File extension without the dot
        
    Returns:
        str: Unique filepath with extension
    """
    # Add extension if not already present
    if not extension.startswith('.'):
        extension = '.' + extension
        
    output_path = base_path + extension
    counter = 1
    
    while os.path.exists(output_path):
        output_path = f"{base_path}_{counter}{extension}"
        counter += 1
        
    return output_path 