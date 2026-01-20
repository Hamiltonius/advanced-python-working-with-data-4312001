#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Renames PDF files based on Excel lookup (single column)

.DESCRIPTION
    Searches for PO numbers (PO_xxxxxxxxxx.pdf) in Excel column and renames PDFs to the full cell text.
    Extracts the 10-digit number from filename, finds it in the Excel column, and renames the PDF
    to the FULL TEXT of that Excel cell.

.PARAMETER PDFFolder
    Path to folder containing PDF files

.PARAMETER ExcelFile
    Path to Excel file with lookup data

.PARAMETER Column
    Column number (1-based) with PO numbers and new names

.PARAMETER SheetName
    Excel sheet name (default: first sheet)

.PARAMETER Execute
    Switch to actually rename files (default is dry-run/preview mode)

.EXAMPLE
    .\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1
    Preview what would be renamed (dry run)

.EXAMPLE
    .\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1 -Execute
    Actually rename the files

.EXAMPLE
    .\Rename-PDFsFromExcel.ps1 -PDFFolder "C:\PDFs" -ExcelFile "C:\lookup.xlsx" -Column 1 -SheetName "Sheet1" -Execute
    Specify sheet name and execute
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$true)]
    [string]$PDFFolder,

    [Parameter(Mandatory=$true)]
    [string]$ExcelFile,

    [Parameter(Mandatory=$true)]
    [int]$Column,

    [Parameter(Mandatory=$false)]
    [string]$SheetName = $null,

    [Parameter(Mandatory=$false)]
    [switch]$Execute
)

# Function to load Excel data
function Load-ExcelData {
    param(
        [string]$ExcelPath,
        [int]$Col,
        [string]$Sheet
    )

    Write-Host "Loading Excel file: $ExcelPath" -ForegroundColor Cyan

    if (-not (Test-Path $ExcelPath)) {
        throw "Excel file not found: $ExcelPath"
    }

    # Create Excel COM object
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

    try {
        $workbook = $excel.Workbooks.Open($ExcelPath)

        # Select sheet
        if ($Sheet) {
            $worksheet = $workbook.Worksheets.Item($Sheet)
        } else {
            $worksheet = $workbook.Worksheets.Item(1)
        }

        Write-Host "Reading sheet: $($worksheet.Name)" -ForegroundColor Cyan

        # Find used range
        $usedRange = $worksheet.UsedRange
        $rowCount = $usedRange.Rows.Count

        # Build lookup hashtable
        $lookupData = @{}
        $loadedCount = 0

        for ($row = 1; $row -le $rowCount; $row++) {
            $cellValue = $worksheet.Cells.Item($row, $Col).Text

            if ($cellValue) {
                # Convert to string and look for 10-digit PO number
                $cellText = $cellValue.Trim()

                # Extract 10-digit number using regex
                if ($cellText -match '\d{10}') {
                    $poNumber = $matches[0]
                    # Store the FULL cell text as the new name
                    $lookupData[$poNumber] = $cellText
                    $loadedCount++
                }
            }
        }

        Write-Host "Loaded $loadedCount lookup entries from Excel" -ForegroundColor Green

        return $lookupData
    }
    finally {
        # Clean up
        $workbook.Close($false)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

# Function to extract PO number from filename
function Get-PONumber {
    param([string]$Filename)

    if ($Filename -match 'PO_(\d{10})') {
        return $matches[1]
    }
    return $null
}

# Main script
try {
    # Validate PDF folder
    if (-not (Test-Path $PDFFolder)) {
        throw "PDF folder not found: $PDFFolder"
    }

    # Load Excel data
    $lookupData = Load-ExcelData -ExcelPath $ExcelFile -Col $Column -Sheet $SheetName

    # Find PDF files matching pattern
    $pdfFiles = Get-ChildItem -Path $PDFFolder -Filter "PO_*.pdf" | Where-Object { $_.Name -match 'PO_\d{10}\.pdf' }

    Write-Host "`nFound $($pdfFiles.Count) PDF files matching pattern PO_xxxxxxxxxx.pdf" -ForegroundColor Cyan

    if ($pdfFiles.Count -eq 0) {
        Write-Host "No PDF files found matching pattern." -ForegroundColor Yellow
        exit 0
    }

    # Counters
    $renamedCount = 0
    $notFoundCount = 0
    $errorCount = 0

    # Header
    Write-Host "`n$('=' * 80)" -ForegroundColor White
    if (-not $Execute) {
        Write-Host "DRY RUN MODE - No files will be renamed" -ForegroundColor Yellow
        Write-Host "Use -Execute switch to actually rename files" -ForegroundColor Yellow
    } else {
        Write-Host "RENAMING FILES" -ForegroundColor Green
    }
    Write-Host "$('=' * 80)`n" -ForegroundColor White

    # Process each PDF
    foreach ($pdf in $pdfFiles) {
        $poNumber = Get-PONumber -Filename $pdf.Name

        if ($poNumber -and $lookupData.ContainsKey($poNumber)) {
            $newName = $lookupData[$poNumber]

            # Ensure new name ends with .pdf
            if (-not $newName.EndsWith('.pdf', [StringComparison]::OrdinalIgnoreCase)) {
                $newName = "$newName.pdf"
            }

            $newPath = Join-Path $pdf.DirectoryName $newName

            # Check if target already exists
            if ((Test-Path $newPath) -and ($newPath -ne $pdf.FullName)) {
                Write-Host "❌ SKIP: $($pdf.Name)" -ForegroundColor Red
                Write-Host "   Target already exists: $newName`n" -ForegroundColor Red
                $errorCount++
                continue
            }

            Write-Host "✓ MATCH: $($pdf.Name)" -ForegroundColor Green
            Write-Host "  → Rename to: $newName" -ForegroundColor Gray

            if ($Execute) {
                try {
                    Rename-Item -Path $pdf.FullName -NewName $newName -ErrorAction Stop
                    Write-Host "  ✓ SUCCESS`n" -ForegroundColor Green
                    $renamedCount++
                }
                catch {
                    Write-Host "  ❌ ERROR: $_`n" -ForegroundColor Red
                    $errorCount++
                }
            }
            else {
                Write-Host ""
                $renamedCount++
            }
        }
        else {
            Write-Host "⚠ NO MATCH: $($pdf.Name)" -ForegroundColor Yellow
            if ($poNumber) {
                Write-Host "  PO Number $poNumber not found in Excel`n" -ForegroundColor Gray
            }
            else {
                Write-Host "  Could not extract PO number`n" -ForegroundColor Gray
            }
            $notFoundCount++
        }
    }

    # Summary
    Write-Host "$('=' * 80)" -ForegroundColor White
    Write-Host "SUMMARY" -ForegroundColor Cyan
    Write-Host "$('=' * 80)" -ForegroundColor White
    if (-not $Execute) {
        Write-Host "Would rename: $renamedCount files" -ForegroundColor Yellow
    } else {
        Write-Host "Successfully renamed: $renamedCount files" -ForegroundColor Green
    }
    Write-Host "Not found in Excel: $notFoundCount files" -ForegroundColor Yellow
    Write-Host "Errors: $errorCount files" -ForegroundColor $(if ($errorCount -gt 0) { 'Red' } else { 'Gray' })
    Write-Host "$('=' * 80)" -ForegroundColor White
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
    exit 1
}
