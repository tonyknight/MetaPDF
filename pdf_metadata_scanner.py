import os
import pandas as pd
from PyPDF2 import PdfReader
from datetime import datetime
import re

def extract_pdf_metadata(filepath):
    """Extract only existing metadata from a PDF file."""
    try:
        reader = PdfReader(filepath)
        
        # Handle encrypted PDFs
        if reader.is_encrypted:
            return create_error_metadata(filepath, 'Encrypted PDF')
        
        # Safely get metadata
        try:
            info = reader.metadata or {}
        except Exception as e:
            return create_error_metadata(filepath, f'Metadata error: {str(e)}')
        
        filename = sanitize_field(os.path.basename(filepath))
        filepath = sanitize_field(filepath)
        
        # Safely extract all metadata fields
        metadata = {}
        for field, key in [
            ('title', '/Title'),
            ('author', '/Author'),
            ('subject', '/Subject'),
            ('tags', '/Keywords'),
        ]:
            try:
                value = info.get(key, None)
                if hasattr(value, 'get_object'):
                    try:
                        value = value.get_object()
                    except Exception:
                        value = None
                metadata[field] = sanitize_field(value) if value else None
            except Exception:
                metadata[field] = None
        
        # Safely extract date
        pdf_date = None
        raw_date = None
        try:
            raw_date = info.get('/CreationDate', None)
            if hasattr(raw_date, 'get_object'):
                try:
                    raw_date = raw_date.get_object()
                except Exception:
                    raw_date = None
                    
            if raw_date and isinstance(raw_date, str) and raw_date.startswith("D:"):
                try:
                    pdf_date = datetime.strptime(raw_date[2:14], "%Y%m%d%H%M")
                except ValueError:
                    pass
        except Exception:
            raw_date = None
        
        return {
            'filename': filename,
            'filepath': filepath,
            'has_title': metadata['title'] is not None,
            'title': metadata['title'],
            'has_author': metadata['author'] is not None,
            'author': metadata['author'],
            'has_subject': metadata['subject'] is not None,
            'subject': metadata['subject'],
            'has_tags': metadata['tags'] is not None,
            'tags': metadata['tags'],
            'has_date': pdf_date is not None,
            'date': pdf_date,
            'raw_date_string': sanitize_field(raw_date),
            'error': None
        }
    except Exception as e:
        error_msg = str(e)
        if "PyCryptodome is required" in error_msg:
            error_msg = "Encrypted PDF (requires PyCryptodome)"
        elif "EOF marker not found" in error_msg:
            error_msg = "Corrupted PDF (EOF marker not found)"
            
        print(f"Error processing {filepath}: {error_msg}")
        return create_error_metadata(filepath, error_msg)

def sanitize_field(value):
    """Replace commas with semicolons in a field value."""
    if value is None:
        return None
    return str(value).replace(',', ';')

def create_error_metadata(filepath, error_msg):
    """Create metadata dictionary for error cases."""
    return {
        'filename': sanitize_field(os.path.basename(filepath)),
        'filepath': sanitize_field(filepath),
        'has_title': False,
        'title': None,
        'has_author': False,
        'author': None,
        'has_subject': False,
        'subject': None,
        'has_tags': False,
        'tags': None,
        'has_date': False,
        'date': None,
        'raw_date_string': None,
        'error': error_msg
    }

def scan_pdfs(root_folder):
    """Recursively scan folder for PDFs and extract metadata."""
    pdf_data = []
    error_data = []
    total_pdfs = 0
    error_counts = {}
    object_error_files = []
    encrypted_files = []
    corrupted_files = []
    
    for root, _, files in os.walk(root_folder):
        for file in files:
            if file.lower().endswith('.pdf'):
                total_pdfs += 1
                filepath = os.path.join(root, file)
                metadata = extract_pdf_metadata(filepath)
                
                if metadata:
                    pdf_data.append(metadata)
                    # Track error types
                    error = metadata.get('error')
                    if error:
                        error_counts[error] = error_counts.get(error, 0) + 1
                        # Track files with specific errors
                        if "Object" in error:
                            object_error_files.append(filepath)
                        if "Encrypted PDF" in error:
                            encrypted_files.append(filepath)
                        if "Corrupted PDF" in error:
                            corrupted_files.append(filepath)
                        # Add to error data
                        error_data.append({
                            'filename': metadata['filename'],
                            'filepath': metadata['filepath'],
                            'error_type': error
                        })

    # Print statistics
    print(f"\nPDF Processing Statistics:")
    print(f"Total PDFs found: {total_pdfs}")
    print(f"Successfully processed (including those with errors): {len(pdf_data)}")
    print(f"Successfully processed (without errors): {len([d for d in pdf_data if not d.get('error')])}")
    
    if error_counts:
        print("\nError Summary:")
        for error_type, count in error_counts.items():
            print(f"- {error_type}: {count} files")
    
    return pdf_data, error_data

def metadata_to_csv():
    """Export PDF metadata to CSV files."""
    # Get current datetime for filenames
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Scan PDFs and extract metadata
    print(f"Starting PDF scan in: {PDF_FOLDER}")
    pdf_data, error_data = scan_pdfs(PDF_FOLDER)
    
    # Create DataFrames and save to CSV files
    if pdf_data:
        # Save main metadata
        df = pd.DataFrame(pdf_data)
        output_file = f"({current_time}) Metadata2CSV.csv"
        df.to_csv(output_file, index=False)
        print(f"\nMetadata saved to {output_file}")
        
        # Save error data if any errors occurred
        if error_data:
            error_df = pd.DataFrame(error_data)
            error_output_file = f"({current_time}) Metadata2CSV Errors.csv"
            error_df.to_csv(error_output_file, index=False)
            print(f"Error data saved to {error_output_file}")
    else:
        print("No PDF files found")

def parse_date_from_parentheses(filename):
    """Extract and parse dates from parenthetical expressions in filename."""
    # First, find all potential date matches
    date_patterns = {
        'full_date': r'[\(\[\{](\d{4}-\d{2}-\d{2})[\)\]\}]',  # (YYYY-MM-DD) or [YYYY-MM-DD] or {YYYY-MM-DD}
        'compact_date': r'[\(\[\{].*?(\d{8}).*?[\)\]\}]',      # (YYYYMMDD) or [YYYYMMDD] or {YYYYMMDD} with possible extra text
        'year_month': r'[\(\[\{](\d{4})[-_\.](\d{2})[\)\]\}]', # (YYYY-MM) or [YYYY-MM] or {YYYY-MM}
        'year': r'[\(\[\{](\d{4})[\)\]\}]',                    # (YYYY) or [YYYY] or {YYYY}
        'text_month_year': r'[\(\[\{]((?:January|February|March|April|May|June|July|August|September|October|November|December))\s+(\d{4})[\)\]\}]', # (Month YYYY)
        'text_month_range': r'[\(\[\{](\d{4})[\)\]\}].*?[\(\[\{]((?:January|February|March|April|May|June|July|August|September|October|November|December))[^\)\]\}]*[\)\]\}]', # (YYYY)...(Month)
        'year_range': r'[\(\[\{](\d{4})-(\d{4})[\)\]\}]',      # (YYYY-YYYY) or [YYYY-YYYY] or {YYYY-YYYY}
        'expense_date': r'[\(\[\{].*?(\d{4}-\d{4})_\d+[\)\]\}]', # (YYYY-MMDD_HHMMSS)
        'short_month': r'[\(\[\{]((?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec))\s+(\d{4})[\)\]\}]', # (MMM YYYY)
    }
    
    month_map = {
        'january': '01', 'february': '02', 'march': '03', 'april': '04',
        'may': '05', 'june': '06', 'july': '07', 'august': '08',
        'september': '09', 'october': '10', 'november': '11', 'december': '12',
        'jan': '01', 'feb': '02', 'mar': '03', 'apr': '04',
        'jun': '06', 'jul': '07', 'aug': '08', 'sep': '09', 'oct': '10', 'nov': '11', 'dec': '12'
    }
    
    dates = []
    year_hint = None
    
    # First pass: Look for full dates and year hints
    for pattern_type in ['full_date', 'compact_date']:
        matches = re.findall(date_patterns[pattern_type], filename, re.IGNORECASE)
        for match in matches:
            try:
                date_str = match if isinstance(match, str) else match[0]
                if len(date_str) == 8 and date_str.isdigit():  # YYYYMMDD
                    formatted_date = f"{date_str[:4]}-{date_str[4:6]}-{date_str[6:8]}"
                else:
                    formatted_date = date_str
                datetime.strptime(formatted_date, "%Y-%m-%d")
                dates.append(formatted_date)
            except ValueError:
                continue
    
    # Look for expense report style dates (YYYY-MMDD)
    matches = re.findall(date_patterns['expense_date'], filename, re.IGNORECASE)
    for match in matches:
        try:
            year, mmdd = match.split('-')
            formatted_date = f"{year}-{mmdd[:2]}-{mmdd[2:4]}"
            datetime.strptime(formatted_date, "%Y-%m-%d")
            dates.append(formatted_date)
        except (ValueError, IndexError):
            continue
    
    # If we found any full dates, use the earliest one
    if dates:
        return min(dates)
    
    # Second pass: Look for year + month combinations
    year_matches = re.findall(date_patterns['year'], filename)
    if year_matches:
        year_hint = year_matches[0]
    
    # Look for text month + year combinations
    for pattern in ['text_month_year', 'short_month']:
        matches = re.findall(date_patterns[pattern], filename, re.IGNORECASE)
        for match in matches:
            month = month_map.get(match[0].lower())
            year = match[1]
            if month and year:
                return f"{year}-{month}-01"
    
    # Look for year + text month combinations
    if year_hint:
        for pattern in ['text_month_range']:
            matches = re.findall(date_patterns[pattern], filename, re.IGNORECASE)
            for match in matches:
                month = month_map.get(match[1].lower())
                if month:
                    return f"{year_hint}-{month}-01"
    
    # Look for year ranges and use the first year
    matches = re.findall(date_patterns['year_range'], filename)
    if matches:
        year = matches[0][0]  # Use the start year
        return f"{year}-06-01"
    
    # If we have a year hint but nothing else, use mid-year
    if year_hint:
        return f"{year_hint}-06-01"
    
    return None

def clean_filename(original_filename):
    """Generate cleaned filename with standardized date format."""
    # Extract the best date from all parenthetical expressions
    date_str = parse_date_from_parentheses(original_filename)
    if not date_str:
        return original_filename, None
    
    # Remove all parenthetical dates from the filename
    cleaned_name = original_filename
    date_patterns = [
        r'[\(\[\{]\d{4}-\d{2}-\d{2}[\)\]\}]',                # (YYYY-MM-DD) or [YYYY-MM-DD] or {YYYY-MM-DD}
        r'[\(\[\{]\d{8}[^\)\]\}]*[\)\]\}]',                  # (YYYYMMDD) with extra text
        r'[\(\[\{]\d{4}[\)\]\}]',                            # (YYYY) or [YYYY] or {YYYY}
        r'[\(\[\{]\d{4}[-_\.]\d{2}[-_\.]\d{2}[\)\]\}]',     # (YYYY-MM-DD) with separators
        r'[\(\[\{]\d{4}-\d{4}[\)\]\}]',                      # (YYYY-YYYY)
        r'[\(\[\{]\d{4}-\d{4}_\d+[\)\]\}]',                  # (YYYY-MMDD_HHMMSS)
        r'[\(\[\{](?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{4}[\)\]\}]', # (MMM YYYY)
        r'[\(\[\{](?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}[\)\]\}]', # (Month YYYY)
        r'[\(\[\{](?:January|February|March|April|May|June|July|August|September|October|November|December)(?:\s*-\s*[A-Za-z]+)?[\)\]\}]',  # {Month - Month}
        r'[\(\[\{]\d{4}[\)\]\}].*?[\(\[\{](?:January|February|March|April|May|June|July|August|September|October|November|December)[^\)\]\}]*[\)\]\}]', # (YYYY)...(Month...)
        r'[\[\{\(]\d{4}\s*-\s*\d{4}[\]\}\)]',               # [YYYY - YYYY] or (YYYY - YYYY)
        r'[\(\[\{]\d{4}[-_\.]\d{2}[\)\]\}]',                # (YYYY-MM)
        r'[\(\[\{][A-Za-z]+\s*-\s*[A-Za-z]+[\)\]\}]',       # {Month - Month} without year
        r'[\(\[\{][A-Za-z]+(?:\s*-\s*[A-Za-z]+)?(?:\s+\d{4})?[\)\]\}]',  # {Month - Month YYYY} or {Month YYYY}
    ]
    
    # First pass: remove date-related patterns
    for pattern in date_patterns:
        cleaned_name = re.sub(pattern, '', cleaned_name, flags=re.IGNORECASE)
    
    # Clean up any double spaces and strip
    cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
    
    # Clean up any empty brackets/parentheses and multiple dashes
    cleaned_name = re.sub(r'[\(\[\{]\s*[\)\]\}]', '', cleaned_name)
    cleaned_name = re.sub(r'-+', '-', cleaned_name)
    
    # Clean up spaces around dashes and parentheses
    cleaned_name = re.sub(r'\s*-\s*', ' - ', cleaned_name)
    cleaned_name = re.sub(r'\(\s+', '(', cleaned_name)
    cleaned_name = re.sub(r'\s+\)', ')', cleaned_name)
    
    # Final cleanup of multiple spaces
    cleaned_name = re.sub(r'\s+', ' ', cleaned_name).strip()
    
    # Remove any trailing/leading dashes after cleanup
    cleaned_name = cleaned_name.strip(' -')
    
    # Prepend the standardized date
    final_name = f"({date_str}) {cleaned_name}"
    
    return final_name, date_str

def clean_dates_dryrun():
    """Preview date cleaning operations without making changes."""
    print("Starting Clean Dates preflight scan...")
    
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    results = []
    files_processed = 0
    files_to_modify = 0
    
    # First, get current metadata
    pdf_data, _ = scan_pdfs(PDF_FOLDER)
    metadata_map = {item['filepath']: item for item in pdf_data}
    
    # Process each PDF file
    for root, _, files in os.walk(PDF_FOLDER):
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue
                
            files_processed += 1
            filepath = os.path.join(root, filename)
            
            # Get original metadata
            metadata = metadata_map.get(filepath, {})
            original_date = metadata.get('date')
            
            # Process filename
            cleaned_filename, new_date_str = clean_filename(filename)
            if cleaned_filename != filename:
                files_to_modify += 1
                
                # Create new datetime with noon time
                if new_date_str:
                    new_date = f"{new_date_str} 12:00:00"
                else:
                    new_date = None
                
                results.append({
                    'filepath': filepath,
                    'original_filename': filename,
                    'corrected_filename': cleaned_filename,
                    'original_pdf_date': original_date,
                    'corrected_pdf_date': new_date
                })
    
    # Save results to CSV
    if results:
        df = pd.DataFrame(results)
        output_file = f"({current_time}) Clean Dates Preflight.csv"
        df.to_csv(output_file, index=False)
        
        print(f"\nClean Dates Preflight Summary:")
        print(f"Total files processed: {files_processed}")
        print(f"Files requiring modification: {files_to_modify}")
        print(f"\nResults saved to: {output_file}")
    else:
        print("\nNo files require date cleaning.")

def clean_dates():
    """Clean up dates in PDF filenames."""
    print("Starting Clean Dates operation...")
    
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    results = []
    files_processed = 0
    files_modified = 0
    errors = []
    
    # Process each PDF file
    for root, _, files in os.walk(PDF_FOLDER):
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue
                
            files_processed += 1
            old_filepath = os.path.join(root, filename)
            
            try:
                # Process filename
                cleaned_filename, new_date_str = clean_filename(filename)
                if cleaned_filename != filename:
                    # Construct new filepath
                    new_filepath = os.path.join(root, cleaned_filename)
                    
                    # Check if destination file already exists
                    if os.path.exists(new_filepath) and old_filepath.lower() != new_filepath.lower():
                        error_msg = f"Cannot rename: {cleaned_filename} already exists"
                        errors.append({
                            'filepath': old_filepath,
                            'original_filename': filename,
                            'intended_filename': cleaned_filename,
                            'error': error_msg
                        })
                        print(f"Error: {error_msg}")
                        continue
                    
                    try:
                        # Rename the file
                        os.rename(old_filepath, new_filepath)
                        files_modified += 1
                        
                        # Record the change
                        results.append({
                            'original_filepath': old_filepath,
                            'original_filename': filename,
                            'new_filepath': new_filepath,
                            'new_filename': cleaned_filename,
                            'date_extracted': new_date_str
                        })
                        
                    except OSError as e:
                        error_msg = f"Failed to rename file: {str(e)}"
                        errors.append({
                            'filepath': old_filepath,
                            'original_filename': filename,
                            'intended_filename': cleaned_filename,
                            'error': error_msg
                        })
                        print(f"Error: {error_msg}")
                        
            except Exception as e:
                error_msg = f"Error processing file: {str(e)}"
                errors.append({
                    'filepath': old_filepath,
                    'original_filename': filename,
                    'intended_filename': None,
                    'error': error_msg
                })
                print(f"Error: {error_msg}")
    
    # Save results to CSV files
    if results:
        # Save successful changes
        df = pd.DataFrame(results)
        output_file = f"({current_time}) Clean Dates Changes.csv"
        df.to_csv(output_file, index=False)
        print(f"\nChanges saved to: {output_file}")
    
    if errors:
        # Save errors
        error_df = pd.DataFrame(errors)
        error_file = f"({current_time}) Clean Dates Errors.csv"
        error_df.to_csv(error_file, index=False)
        print(f"Errors saved to: {error_file}")
    
    # Print summary
    print(f"\nClean Dates Summary:")
    print(f"Total files processed: {files_processed}")
    print(f"Files successfully renamed: {files_modified}")
    print(f"Errors encountered: {len(errors)}")

def preview_scan():
    """Preview metadata changes before writing."""
    print("Preview Scan - Not implemented yet")

def metadata_write_dryrun():
    """Preview metadata write operations without making changes."""
    print("Metadata Write Dryrun - Not implemented yet")

def metadata_write():
    """Write metadata to PDF files."""
    print("Metadata Write - Not implemented yet")

def display_menu():
    """Display the main menu."""
    print("\nPDF Metadata Tools")
    print("=================")
    print("1. Metadata2CSV")
    print("2. Clean Dates Dryrun")
    print("3. Clean Dates")
    print("4. Outlier Scan")
    print("5. Metadata Write Dryrun")
    print("6. Metadata Write")
    print("Q. Quit Script")
    return input("\nSelect an option: ").strip().upper()

def main():
    # Configuration
    global PDF_FOLDER
    PDF_FOLDER = "/Users/knight/Desktop/PDFs"
    
    while True:
        try:
            choice = display_menu()
            
            if choice == 'Q':
                print("Exiting script...")
                break
            
            # Process menu choice
            if choice == '1':
                metadata_to_csv()
            elif choice == '2':
                clean_dates_dryrun()
            elif choice == '3':
                clean_dates()
            elif choice == '4':
                outlier_scan()
            elif choice == '5':
                metadata_write_dryrun()
            elif choice == '6':
                metadata_write()
            else:
                print("Invalid option. Please try again.")
            
            # Pause before showing menu again
            input("\nPress Enter to continue...")
            
        except Exception as e:
            print(f"\nAn error occurred: {str(e)}")
            input("\nPress Enter to continue...")

def find_embedded_dates(filename, existing_date=None):
    """Find dates in filename that aren't at the start."""
    # Remove the existing prepended date if present
    if existing_date:
        working_name = filename.replace(f"({existing_date})", "").strip()
    else:
        # Try to find and remove any leading date
        match = re.match(r'\(\d{4}(?:-\d{2}){0,2}\)\s*(.+)', filename)
        working_name = match.group(1) if match else filename
    
    # Store the original year prefix if it exists
    year_prefix_match = re.match(r'\((\d{4})\)\s*(.+)', filename)
    year_prefix = f"({year_prefix_match.group(1)})" if year_prefix_match else None
    
    # Patterns for finding embedded dates
    date_patterns = [
        # Spaced date after prefix (1991 - 01 - 23)
        (r'(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})\s*-\s*',  # 1991 - 01 - 23 -
         lambda m: (f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}", 
                   [f"{m.group(0)}"])),  # Include the trailing dash and space
        
        # Expense report timestamp pattern
        (r'\s*\(\d{4}-\d{2}\d{2}_\d+\)',                    # (2010-0805_162655)
         lambda m: (f"{m.group(0)[2:6]}-{m.group(0)[7:9]}-{m.group(0)[9:11]}", 
                   [m.group(0), year_prefix] if year_prefix else [m.group(0)])),
        
        # Other spaced date format
        (r'(?:^|\s+)(\d{4})\s*-\s*(\d{1,2})\s*-\s*(\d{1,2})(?:\s+|$)',  # 1991 - 01 - 23
         lambda m: (f"{m.group(1)}-{m.group(2).zfill(2)}-{m.group(3).zfill(2)}", [m.group(0)])),
    ]
    
    for pattern, formatter in date_patterns:
        matches = re.finditer(pattern, working_name)
        for match in matches:
            try:
                result = formatter(match)
                if not result:
                    continue
                    
                formatted_date, text_to_remove = result
                
                # Validate the date
                datetime.strptime(formatted_date, "%Y-%m-%d")
                
                # Only return if this is a different/better date than what we already have
                if not existing_date or formatted_date != existing_date:
                    return formatted_date, text_to_remove
                
            except (ValueError, IndexError):
                continue
    
    return None, None

def clean_trailing_separators(filename):
    """Clean up trailing spaces and separators in filename."""
    # Split filename and extension
    name, ext = os.path.splitext(filename)
    
    # Clean up multiple spaces
    name = re.sub(r'\s+', ' ', name)
    
    # Clean up trailing spaces and separators
    name = name.rstrip()
    name = re.sub(r'\s*-\s*$', '', name)
    
    return f"{name}{ext}"

def outlier_scan():
    """Scan for and clean up filename outliers and embedded dates."""
    print("Starting Outlier Scan...")
    
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    separator_cleanups = []
    date_cleanups = []
    errors = []
    files_processed = 0
    
    for root, _, files in os.walk(PDF_FOLDER):
        for filename in files:
            if not filename.lower().endswith('.pdf'):
                continue
                
            files_processed += 1
            filepath = os.path.join(root, filename)
            
            try:
                # Step 1: Clean trailing separators
                cleaned_filename = clean_trailing_separators(filename)
                if cleaned_filename != filename:
                    new_filepath = os.path.join(root, cleaned_filename)
                    
                    # Check if destination file already exists
                    if os.path.exists(new_filepath) and filepath.lower() != new_filepath.lower():
                        error_msg = f"Cannot rename: {cleaned_filename} already exists"
                        errors.append({
                            'filepath': filepath,
                            'original_filename': filename,
                            'intended_filename': cleaned_filename,
                            'error': error_msg
                        })
                        print(f"Error: {error_msg}")
                        continue
                    
                    try:
                        # Rename the file
                        os.rename(filepath, new_filepath)
                        separator_cleanups.append({
                            'filepath': filepath,
                            'original_filename': filename,
                            'cleaned_filename': cleaned_filename,
                            'cleanup_type': 'trailing_separator'
                        })
                        # Update filepath for next step
                        filepath = new_filepath
                        filename = cleaned_filename
                    except OSError as e:
                        error_msg = f"Failed to rename file: {str(e)}"
                        errors.append({
                            'filepath': filepath,
                            'original_filename': filename,
                            'intended_filename': cleaned_filename,
                            'error': error_msg
                        })
                        print(f"Error: {error_msg}")
                
                # Step 2: Look for embedded dates
                match = re.match(r'\((\d{4}(?:-\d{2}){0,2})\)', filename)
                existing_date = match.group(1) if match else None
                
                embedded_date, text_to_remove = find_embedded_dates(filename, existing_date)
                if embedded_date and text_to_remove:
                    # Remove all specified text patterns and clean up the filename
                    working_name = filename
                    for text in text_to_remove:
                        working_name = working_name.replace(text, '')
                    working_name = clean_trailing_separators(working_name)
                    
                    # If we found a better date than existing, propose the change
                    if not existing_date or (len(embedded_date) > len(existing_date)):
                        proposed_filename = f"({embedded_date}) {working_name}"
                        
                        print(f"\nFound potential embedded date:")
                        print(f"Original: {filename}")
                        print(f"Proposed: {proposed_filename}")
                        
                        choice = input("Accept this change? (y/n): ").strip().lower()
                        if choice == 'y':
                            new_filepath = os.path.join(root, proposed_filename)
                            
                            # Check if destination file exists
                            if os.path.exists(new_filepath) and filepath.lower() != new_filepath.lower():
                                error_msg = f"Cannot rename: {proposed_filename} already exists"
                                errors.append({
                                    'filepath': filepath,
                                    'original_filename': filename,
                                    'intended_filename': proposed_filename,
                                    'error': error_msg
                                })
                                print(f"Error: {error_msg}")
                                continue
                            
                            try:
                                # Rename the file
                                os.rename(filepath, new_filepath)
                                date_cleanups.append({
                                    'filepath': filepath,
                                    'original_filename': filename,
                                    'cleaned_filename': proposed_filename,
                                    'original_date': existing_date,
                                    'new_date': embedded_date,
                                    'cleanup_type': 'embedded_date'
                                })
                            except OSError as e:
                                error_msg = f"Failed to rename file: {str(e)}"
                                errors.append({
                                    'filepath': filepath,
                                    'original_filename': filename,
                                    'intended_filename': proposed_filename,
                                    'error': error_msg
                                })
                                print(f"Error: {error_msg}")
                
            except Exception as e:
                error_msg = f"Error processing file: {str(e)}"
                errors.append({
                    'filepath': filepath,
                    'original_filename': filename,
                    'intended_filename': None,
                    'error': error_msg
                })
                print(f"Error: {error_msg}")
    
    # Save results if any changes were found
    if separator_cleanups or date_cleanups:
        all_cleanups = separator_cleanups + date_cleanups
        df = pd.DataFrame(all_cleanups)
        output_file = f"({current_time}) Outlier Scan Results.csv"
        df.to_csv(output_file, index=False)
        print(f"\nResults saved to: {output_file}")
    
    if errors:
        error_df = pd.DataFrame(errors)
        error_file = f"({current_time}) Outlier Scan Errors.csv"
        error_df.to_csv(error_file, index=False)
        print(f"Errors saved to: {error_file}")
    
    # Print summary
    print(f"\nOutlier Scan Summary:")
    print(f"Total files processed: {files_processed}")
    print(f"Files with trailing separators cleaned: {len(separator_cleanups)}")
    print(f"Files with embedded dates identified and accepted: {len(date_cleanups)}")
    print(f"Errors encountered: {len(errors)}")

if __name__ == "__main__":
    main() 