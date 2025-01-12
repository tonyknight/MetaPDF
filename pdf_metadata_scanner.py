import os
import pandas as pd
from PyPDF2 import PdfReader
from datetime import datetime

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

def main():
    # Replace with your PDF folder path
    pdf_folder = "/Users/knight/Sync/Documents/PDFs"
    
    # Get current datetime for filenames
    current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Scan PDFs and extract metadata
    print(f"Starting PDF scan in: {pdf_folder}")
    pdf_data, error_data = scan_pdfs(pdf_folder)
    
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

if __name__ == "__main__":
    main() 