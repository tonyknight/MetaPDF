# PDF Metadata Manager

A Python script for managing and standardizing PDF metadata and filenames. This tool helps organize PDF collections by standardizing dates, cleaning filenames, and managing PDF metadata tags.

## Features

### 1. Metadata to CSV (Option 1)
Scans all PDFs in the specified directory and exports their current metadata to CSV files:
- Creates `Metadata2CSV.csv` with complete metadata inventory
- Creates `Metadata2CSV Errors.csv` for files with processing errors
- Reports statistics on processed files and encountered errors

### 2. Clean Dates Dryrun (Option 2)
Previews date standardization changes without modifying files:
- Identifies dates in various formats within filenames
- Shows proposed filename changes with standardized date format
- Creates `Clean Dates Preflight.csv` showing all proposed changes

### 3. Clean Dates (Option 3)
Performs the actual date standardization:
- Extracts and standardizes dates from filenames
- Moves standardized dates to the start of filename in (YYYY-MM-DD) format
- Creates `Clean Dates Results.csv` documenting all changes
- Handles various date formats:
  - Full dates (YYYY-MM-DD)
  - Year only (YYYY) → converts to mid-year (YYYY-06-01)
  - Month/Year combinations
  - Text month formats
  - Expense report timestamps

### 4. Outlier Scan (Option 4)
Identifies and cleans up filename anomalies:
- Removes trailing spaces and separators
- Identifies embedded dates in filenames
- Allows user to approve date standardization changes
- Creates `Outlier Scan Results.csv` documenting changes
- Creates `Outlier Scan Errors.csv` for failed operations

### 5. Metadata Write Dryrun (Option 5)
Previews metadata changes before applying them:
- Shows which fields will be written to each PDF
- Identifies files that will receive new metadata
- Creates preflight report of all proposed changes

### 6. Metadata Write (Option 6)
Writes metadata to PDF files based on filename parsing:
- DateTime: Always writes standardized date from filename
- Author: Writes first parenthetical value after date if PDF lacks author
- Tags: Writes content from square brackets as PDF keywords
- Title/Subject: Uses cleaned filename (minus date) if fields are empty
- Creates three CSV files:
  - `PDF Metadata.csv`: Successfully written metadata
  - `Untagged.csv`: Files with only date metadata
  - `Skipped Files.csv`: Files that couldn't be processed

### 7. Clean Metadata Fields (Option 7)
Cleans up existing PDF metadata fields:
- Removes '.pdf' extension from Title/Subject fields
- Removes trailing spaces and dashes
- Standardizes spacing
- Creates `Clean.csv` documenting all changes

## Filename Structure

The script parses filenames that follow a specific structure to extract metadata:

1. **Date**: Appears first in parentheses - (YYYY-MM-DD) or (YYYY)
2. **Author**: First parenthetical value after the date - (Author Name)
3. **Tags**: Any text in square brackets - [Tag1][Tag2]
4. **Title**: Remaining text (excluding date, author, and tags)

Examples:

## Usage

1. Set the PDF_FOLDER variable to your target directory
2. Run the script
3. Choose options from the menu to process your PDFs
4. Review CSV output files for results and any errors

## Requirements

- Python 3.x
- PyPDF2
- pandas

## Notes

- Always backup your PDFs before running metadata modifications
- Some PDFs may be encrypted or have other restrictions preventing metadata changes
- The script creates detailed logs of all operations for review
- Each output file is timestamped for tracking multiple runs
