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
            return {
                'filename': os.path.basename(filepath),
                'filepath': filepath,
                'has_title': False,
                'title_if_exists': None,
                'has_date': False,
                'date_if_exists': None,
                'raw_date_string': None,
                'error': 'Encrypted PDF'
            }
        
        # Safely get metadata
        try:
            info = reader.metadata or {}
        except Exception as e:
            return {
                'filename': os.path.basename(filepath),
                'filepath': filepath,
                'has_title': False,
                'title_if_exists': None,
                'has_date': False,
                'date_if_exists': None,
                'raw_date_string': None,
                'error': f'Metadata error: {str(e)}'
            }
        
        filename = os.path.basename(filepath)
        
        # Safely extract title
        try:
            pdf_title = info.get('/Title', None)
            if hasattr(pdf_title, 'get_object'):
                try:
                    pdf_title = pdf_title.get_object()
                except Exception:
                    pdf_title = None
        except Exception:
            pdf_title = None
        
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
            'has_title': pdf_title is not None,
            'title_if_exists': pdf_title,
            'has_date': pdf_date is not None,
            'date_if_exists': pdf_date,
            'raw_date_string': raw_date,
            'error': None
        }
    except Exception as e:
        error_msg = str(e)
        if "PyCryptodome is required" in error_msg:
            error_msg = "Encrypted PDF (requires PyCryptodome)"
        elif "EOF marker not found" in error_msg:
            error_msg = "Corrupted PDF (EOF marker not found)"
            
        print(f"Error processing {filepath}: {error_msg}")
        return {
            'filename': os.path.basename(filepath),
            'filepath': filepath,
            'has_title': False,
            'title_if_exists': None,
            'has_date': False,
            'date_if_exists': None,
            'raw_date_string': None,
            'error': error_msg
        }

def scan_pdfs(root_folder):
    """Recursively scan folder for PDFs and extract metadata."""
    pdf_data = []
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

    # Print statistics
    print(f"\nPDF Processing Statistics:")
    print(f"Total PDFs found: {total_pdfs}")
    print(f"Successfully processed (including those with errors): {len(pdf_data)}")
    print(f"Successfully processed (without errors): {len([d for d in pdf_data if not d.get('error')])}")
    
    if error_counts:
        print("\nError Summary:")
        for error_type, count in error_counts.items():
            print(f"- {error_type}: {count} files")
    
    if encrypted_files:
        print("\nEncrypted PDFs:")
        for file in encrypted_files:
            print(f"- {file}")
    
    if corrupted_files:
        print("\nCorrupted PDFs (EOF marker not found):")
        for file in corrupted_files:
            print(f"- {file}")
            
    if object_error_files:
        print("\nFiles with Object errors:")
        for file in object_error_files:
            print(f"- {file}")
    
    return pdf_data

def main():
    # Replace with your PDF folder path
    pdf_folder = "/Users/knight/Sync/Documents/PDFs"
    
    # Scan PDFs and extract metadata
    print(f"Starting PDF scan in: {pdf_folder}")
    pdf_data = scan_pdfs(pdf_folder)
    
    # Create DataFrame and save to CSV
    if pdf_data:
        df = pd.DataFrame(pdf_data)
        output_file = "pdf_metadata.csv"
        df.to_csv(output_file, index=False)
        print(f"\nMetadata saved to {output_file}")
    else:
        print("No PDF files found")

if __name__ == "__main__":
    main() 