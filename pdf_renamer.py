#!/usr/bin/env python3
"""
PDF Renamer - Renames PDF files based on Excel lookup
Searches for PO numbers in Excel and renames PDFs accordingly
"""

import os
import re
import sys
from pathlib import Path
from openpyxl import load_workbook


class PDFRenamer:
    def __init__(self, pdf_folder, excel_file, search_column, rename_column, sheet_name=None):
        """
        Initialize PDF Renamer

        Args:
            pdf_folder: Path to folder containing PDF files
            excel_file: Path to Excel file
            search_column: Column letter/number to search for PO numbers (e.g., 'A' or 1)
            rename_column: Column letter/number with new names (e.g., 'B' or 2)
            sheet_name: Excel sheet name (uses first sheet if None)
        """
        self.pdf_folder = Path(pdf_folder)
        self.excel_file = Path(excel_file)
        self.search_column = search_column
        self.rename_column = rename_column
        self.sheet_name = sheet_name
        self.lookup_data = {}

    def load_excel_data(self):
        """Load Excel data and create lookup dictionary"""
        print(f"Loading Excel file: {self.excel_file}")

        if not self.excel_file.exists():
            raise FileNotFoundError(f"Excel file not found: {self.excel_file}")

        workbook = load_workbook(self.excel_file, read_only=True, data_only=True)

        # Use specified sheet or first sheet
        if self.sheet_name:
            sheet = workbook[self.sheet_name]
        else:
            sheet = workbook.active

        print(f"Reading sheet: {sheet.title}")

        # Convert column letters to numbers if needed
        search_col = self._get_column_number(self.search_column)
        rename_col = self._get_column_number(self.rename_column)

        # Build lookup dictionary
        row_count = 0
        for row in sheet.iter_rows(min_row=1, values_only=True):
            search_value = row[search_col - 1] if len(row) >= search_col else None
            rename_value = row[rename_col - 1] if len(row) >= rename_col else None

            if search_value and rename_value:
                # Extract 10-digit number from search value if present
                search_str = str(search_value)
                po_match = re.search(r'\d{10}', search_str)
                if po_match:
                    po_number = po_match.group()
                    self.lookup_data[po_number] = str(rename_value).strip()
                    row_count += 1

        workbook.close()
        print(f"Loaded {row_count} lookup entries from Excel")
        return row_count

    def _get_column_number(self, col):
        """Convert column letter to number or return number if already numeric"""
        if isinstance(col, int):
            return col
        if isinstance(col, str):
            if col.isdigit():
                return int(col)
            # Convert letter to number (A=1, B=2, etc.)
            col = col.upper()
            num = 0
            for char in col:
                num = num * 26 + (ord(char) - ord('A') + 1)
            return num
        raise ValueError(f"Invalid column specification: {col}")

    def extract_po_number(self, filename):
        """Extract 10-digit PO number from filename"""
        match = re.search(r'PO_(\d{10})', filename)
        if match:
            return match.group(1)
        return None

    def find_pdfs(self):
        """Find all PDF files matching PO_xxxxxxxxxx pattern"""
        pdf_files = []
        pattern = re.compile(r'PO_\d{10}\.pdf', re.IGNORECASE)

        for file_path in self.pdf_folder.glob('*.pdf'):
            if pattern.match(file_path.name):
                pdf_files.append(file_path)

        print(f"Found {len(pdf_files)} PDF files matching pattern PO_xxxxxxxxxx.pdf")
        return pdf_files

    def rename_pdfs(self, dry_run=True):
        """
        Rename PDF files based on Excel lookup

        Args:
            dry_run: If True, only show what would be renamed without actually renaming
        """
        pdf_files = self.find_pdfs()

        if not pdf_files:
            print("No PDF files found matching pattern PO_xxxxxxxxxx.pdf")
            return

        renamed_count = 0
        not_found_count = 0
        error_count = 0

        print("\n" + "="*80)
        if dry_run:
            print("DRY RUN MODE - No files will be renamed")
        else:
            print("RENAMING FILES")
        print("="*80 + "\n")

        for pdf_path in pdf_files:
            po_number = self.extract_po_number(pdf_path.name)

            if po_number and po_number in self.lookup_data:
                new_name = self.lookup_data[po_number]

                # Ensure new name ends with .pdf
                if not new_name.lower().endswith('.pdf'):
                    new_name = f"{new_name}.pdf"

                new_path = pdf_path.parent / new_name

                # Check if target file already exists
                if new_path.exists() and new_path != pdf_path:
                    print(f"❌ SKIP: {pdf_path.name}")
                    print(f"   Target already exists: {new_name}\n")
                    error_count += 1
                    continue

                print(f"✓ MATCH: {pdf_path.name}")
                print(f"  → Rename to: {new_name}")

                if not dry_run:
                    try:
                        pdf_path.rename(new_path)
                        print(f"  ✓ SUCCESS\n")
                        renamed_count += 1
                    except Exception as e:
                        print(f"  ❌ ERROR: {e}\n")
                        error_count += 1
                else:
                    print()
                    renamed_count += 1
            else:
                print(f"⚠ NO MATCH: {pdf_path.name}")
                if po_number:
                    print(f"  PO Number {po_number} not found in Excel\n")
                else:
                    print(f"  Could not extract PO number\n")
                not_found_count += 1

        # Summary
        print("="*80)
        print("SUMMARY")
        print("="*80)
        if dry_run:
            print(f"Would rename: {renamed_count} files")
        else:
            print(f"Successfully renamed: {renamed_count} files")
        print(f"Not found in Excel: {not_found_count} files")
        print(f"Errors: {error_count} files")
        print("="*80)


def main():
    """Main entry point"""
    import argparse

    parser = argparse.ArgumentParser(
        description='Rename PDF files based on Excel lookup',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Dry run (preview changes)
  python pdf_renamer.py /path/to/pdfs lookup.xlsx A B

  # Actually rename files
  python pdf_renamer.py /path/to/pdfs lookup.xlsx A B --execute

  # Specify sheet name
  python pdf_renamer.py /path/to/pdfs lookup.xlsx A B --sheet "Sheet1" --execute

  # Use column numbers instead of letters
  python pdf_renamer.py /path/to/pdfs lookup.xlsx 1 2 --execute
        """
    )

    parser.add_argument('pdf_folder', help='Folder containing PDF files')
    parser.add_argument('excel_file', help='Excel file with lookup data')
    parser.add_argument('search_column', help='Column to search for PO numbers (letter or number, e.g., A or 1)')
    parser.add_argument('rename_column', help='Column with new filenames (letter or number, e.g., B or 2)')
    parser.add_argument('--sheet', help='Excel sheet name (default: first sheet)')
    parser.add_argument('--execute', action='store_true', help='Actually rename files (default is dry-run)')

    args = parser.parse_args()

    try:
        renamer = PDFRenamer(
            args.pdf_folder,
            args.excel_file,
            args.search_column,
            args.rename_column,
            args.sheet
        )

        renamer.load_excel_data()
        renamer.rename_pdfs(dry_run=not args.execute)

    except Exception as e:
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()
