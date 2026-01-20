# PDF Renamer - Excel Lookup Based Renaming

Automatically rename PDF files based on Excel lookup data. Searches for PO numbers in your PDFs (format: `PO_xxxxxxxxxx.pdf`) and renames them to the full text found in your Excel column.

## Features

- ✅ Extracts 10-digit PO numbers from PDF filenames
- ✅ Looks up PO numbers in Excel column
- ✅ Renames PDFs to the FULL TEXT of the matching Excel cell
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
python pdf_renamer.py /path/to/pdfs lookup.xlsx A
```

**Actually rename files:**
```bash
python pdf_renamer.py /path/to/pdfs lookup.xlsx A --execute
```

**With specific sheet:**
```bash
python pdf_renamer.py /path/to/pdfs lookup.xlsx A --sheet "Sheet1" --execute
```

### Option 2: PowerShell Version (Windows)

**Preview changes (dry run):**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1
```

**Actually rename files:**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1 -Execute
```

**With specific sheet:**
```powershell
.\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1 -SheetName "Sheet1" -Execute
```

## How It Works

1. **Scans PDF folder** for files matching pattern `PO_xxxxxxxxxx.pdf`
2. **Extracts 10-digit number** from each PDF filename
3. **Searches Excel column** for that number
4. **When match found**, renames PDF to the **FULL TEXT** of that Excel cell
5. **Reports** all matches, non-matches, and errors

## Example

**Excel file (lookup.xlsx), Column A:**
| Column A |
|----------|
| 1234567890 - Invoice ABC Corp |
| 9876543210 - Receipt XYZ Inc |
| 5555555555 - Contract ACME Ltd |

**PDF files before:**
```
PO_1234567890.pdf
PO_9876543210.pdf
PO_5555555555.pdf
PO_1111111111.pdf  (not in Excel)
```

**PDF files after execution:**
```
1234567890 - Invoice ABC Corp.pdf
9876543210 - Receipt XYZ Inc.pdf
5555555555 - Contract ACME Ltd.pdf
PO_1111111111.pdf  (unchanged - no match)
```

## Parameters

### Python Version
- `pdf_folder` - Folder containing PDF files (required)
- `excel_file` - Excel file with lookup data (required)
- `column` - Column with PO numbers and new names, e.g., 'A' or '1' (required)
- `--sheet` - Excel sheet name (optional, default: first sheet)
- `--execute` - Actually rename files (optional, default: dry-run mode)

### PowerShell Version
- `-PDFFolder` - Folder containing PDF files (required)
- `-ExcelFile` - Excel file with lookup data (required)
- `-Column` - Column number with PO numbers and new names, e.g., 1 (required)
- `-SheetName` - Excel sheet name (optional, default: first sheet)
- `-Execute` - Switch to actually rename files (optional, default: dry-run mode)

## Safety Features

- **Dry-run by default** - Preview changes before executing
- **Duplicate detection** - Won't overwrite existing files
- **Error handling** - Reports all errors clearly
- **Clear reporting** - Shows exactly what will be renamed
- **PO number extraction** - Finds 10-digit numbers anywhere in Excel cells

## Common Use Cases

### Test before executing (for ~50 files)
```bash
# Preview what would happen
python pdf_renamer.py ./invoices data.xlsx A

# If it looks good, execute
python pdf_renamer.py ./invoices data.xlsx A --execute
```

### Column letters vs numbers
```bash
# Using letters (Python only)
python pdf_renamer.py ./pdfs data.xlsx A --execute

# Using numbers (both Python and PowerShell)
python pdf_renamer.py ./pdfs data.xlsx 1 --execute
```

### Multiple sheets
```bash
# Specify sheet name
python pdf_renamer.py ./pdfs data.xlsx A --sheet "January" --execute
```

## Troubleshooting

**No PDFs found:**
- Check that PDF filenames match pattern `PO_xxxxxxxxxx.pdf` (exactly 10 digits)
- Verify the PDF folder path is correct

**No matches in Excel:**
- Verify the Excel column contains cells with 10-digit numbers
- Check that Excel file path is correct
- Make sure you're using the right sheet name
- The 10-digit number can appear ANYWHERE in the cell text

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
- The 10-digit number can appear ANYWHERE in the Excel cell (beginning, middle, end)
- The ENTIRE cell text becomes the new PDF filename
- New filenames will automatically get `.pdf` extension if missing
- Both scripts are safe to run multiple times
- **Always run in dry-run mode first to preview changes**
- Perfect for bulk renaming ~50 files at once

## Support

If you encounter issues:
1. **Run in dry-run mode first** (without `--execute` or `-Execute`)
2. Check file paths and permissions
3. Verify Excel column specification
4. Check that PO numbers match exactly (10 digits)
5. Ensure Excel cells contain both the PO number and descriptive text
