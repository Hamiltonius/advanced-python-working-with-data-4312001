# PDF Renamer - Excel Lookup Based Renaming

Automatically rename PDF files based on Excel lookup data. Searches for PO numbers in your PDFs (format: `PO_xxxxxxxxxx.pdf`) and renames them based on matching values in an Excel spreadsheet.

## Features

- ✅ Extracts 10-digit PO numbers from PDF filenames
- ✅ Looks up PO numbers in Excel column
- ✅ Renames PDFs based on corresponding Excel column value
- ✅ Dry-run mode (preview changes before executing)
- ✅ Safe operation (checks for existing files)
- ✅ Detailed progress and summary reports
- ✅ Two versions: Python and PowerShell

## Quick Start

### Option 1: Python Version (Recommended - Cross-platform)

**Install dependencies:**
```bash
pip install -r requirements_pdf_renamer.txt
```

**Preview changes (dry run):**
```bash
python pdf_renamer.py /path/to/pdfs lookup.xlsx A B
```

**Actually rename files:**
```bash
python pdf_renamer.py /path/to/pdfs lookup.xlsx A B --execute
```

**With specific sheet:**
```bash
python pdf_renamer.py /path/to/pdfs lookup.xlsx A B --sheet "Sheet1" --execute
```

### Option 2: PowerShell Version (Windows)

**Preview changes (dry run):**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -SearchColumn 1 -RenameColumn 2
```

**Actually rename files:**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -SearchColumn 1 -RenameColumn 2 -Execute
```

**With specific sheet:**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -SearchColumn 1 -RenameColumn 2 -SheetName "Sheet1" -Execute
```

## How It Works

1. **Scans PDF folder** for files matching pattern `PO_xxxxxxxxxx.pdf`
2. **Extracts 10-digit number** from each PDF filename
3. **Searches Excel column** for that number
4. **When match found**, renames PDF to the value in the corresponding rename column
5. **Reports** all matches, non-matches, and errors

## Example

**Excel file (lookup.xlsx):**
| Column A (PO Number) | Column B (New Name) |
|---------------------|---------------------|
| 1234567890 | Invoice_ABC_Corp.pdf |
| 9876543210 | Receipt_XYZ_Inc.pdf |
| 5555555555 | Contract_ACME.pdf |

**PDF files before:**
```
PO_1234567890.pdf
PO_9876543210.pdf
PO_5555555555.pdf
PO_1111111111.pdf  (not in Excel)
```

**PDF files after execution:**
```
Invoice_ABC_Corp.pdf
Receipt_XYZ_Inc.pdf
Contract_ACME.pdf
PO_1111111111.pdf  (unchanged - no match)
```

## Parameters

### Python Version
- `pdf_folder` - Folder containing PDF files (required)
- `excel_file` - Excel file with lookup data (required)
- `search_column` - Column to search for PO numbers, e.g., 'A' or '1' (required)
- `rename_column` - Column with new filenames, e.g., 'B' or '2' (required)
- `--sheet` - Excel sheet name (optional, default: first sheet)
- `--execute` - Actually rename files (optional, default: dry-run mode)

### PowerShell Version
- `-PDFFolder` - Folder containing PDF files (required)
- `-ExcelFile` - Excel file with lookup data (required)
- `-SearchColumn` - Column number to search for PO numbers, e.g., 1 (required)
- `-RenameColumn` - Column number with new filenames, e.g., 2 (required)
- `-SheetName` - Excel sheet name (optional, default: first sheet)
- `-Execute` - Switch to actually rename files (optional, default: dry-run mode)

## Safety Features

- **Dry-run by default** - Preview changes before executing
- **Duplicate detection** - Won't overwrite existing files
- **Error handling** - Reports all errors clearly
- **Clear reporting** - Shows exactly what will be renamed
- **PO number extraction** - Finds 10-digit numbers anywhere in search column

## Common Use Cases

### Test before executing
```bash
# Preview what would happen
python pdf_renamer.py ./invoices data.xlsx 1 3

# If it looks good, execute
python pdf_renamer.py ./invoices data.xlsx 1 3 --execute
```

### Column letters vs numbers
```bash
# Using letters (Python only)
python pdf_renamer.py ./pdfs data.xlsx A B --execute

# Using numbers (both Python and PowerShell)
python pdf_renamer.py ./pdfs data.xlsx 1 2 --execute
```

### Multiple sheets
```bash
# Specify sheet name
python pdf_renamer.py ./pdfs data.xlsx A B --sheet "January" --execute
```

## Troubleshooting

**No PDFs found:**
- Check that PDF filenames match pattern `PO_xxxxxxxxxx.pdf` (exactly 10 digits)
- Verify the PDF folder path is correct

**No matches in Excel:**
- Verify the search column contains the 10-digit PO numbers
- Check that Excel file path is correct
- Make sure you're using the right sheet name

**Permission errors:**
- Ensure you have write permissions in the PDF folder
- Close any PDFs that are open
- Run as administrator (Windows) if needed

**Python: openpyxl not found:**
```bash
pip install openpyxl
```

**PowerShell: Excel COM errors:**
- Ensure Microsoft Excel is installed on Windows
- Close any open Excel files

## Requirements

### Python Version
- Python 3.6+
- openpyxl library (`pip install openpyxl`)

### PowerShell Version
- Windows with PowerShell 5.1+
- Microsoft Excel installed (uses COM automation)

## Notes

- PDF filenames must match pattern: `PO_` followed by exactly 10 digits
- The 10-digit number can appear anywhere in the Excel search column
- New filenames from Excel will automatically get `.pdf` extension if missing
- Both scripts are safe to run multiple times
- Always run in dry-run mode first to preview changes

## Support

If you encounter issues:
1. Run in dry-run mode first (without `--execute`)
2. Check file paths and permissions
3. Verify Excel column specifications
4. Check that PO numbers match exactly (10 digits)
