# EDF GDC PICKS Email Automation

Automatically extract table data from daily EDF GDC PICKS emails and populate your Excel workbook.

## Features

- **Automatic Email Detection**: Finds all emails with "EDF GDC PICKS - MM/DD/YYYY" in the subject
- **Table Extraction**: Parses HTML tables from email bodies
- **Chamber Identification**: Automatically identifies Dry, Frozen Dairy Deli, and Meat & Produce data
- **Excel Integration**: Updates your FY26 EDF Data workbook automatically
- **Date Extraction**: Uses the date from the email subject for the Date column
- **Email Cleanup**: Deletes processed emails (optional)
- **Dry Run Mode**: Test the script safely before applying changes

## Project Structure

```
edf-email-automation/
├── .venv/                    # Virtual environment
├── main.py                   # Main orchestration script
├── outlook_handler.py        # Outlook COM interface
├── table_parser.py          # HTML table parsing
├── excel_handler.py         # Excel workbook manipulation
├── run_live.bat             # Quick run script (live mode)
├── run_dry_run.bat          # Quick run script (dry run mode)
└── README.md                # This file
```

## Prerequisites

- Windows 10/11
- Microsoft Outlook (Desktop version)
- Python 3.10+
- Virtual environment with required packages (already set up)

## Installation

The project is already set up with all dependencies installed:

```bash
# Dependencies installed:
# - openpyxl (Excel manipulation)
# - beautifulsoup4 (HTML parsing)
# - pywin32 (Outlook COM interface)
# - python-dateutil (Date utilities)
```

## Usage

### Option 1: Using Batch Files (Easiest)

**Dry Run (Test without making changes):**
```bash
run_dry_run.bat
```

**Live Mode (Actually updates Excel and deletes emails):**
```bash
run_live.bat
```

### Option 2: Command Line

**Dry Run:**
```bash
.venv\Scripts\python.exe main.py
```

**Live Mode:**
```bash
.venv\Scripts\python.exe main.py --live
```

## Workflow

1. **Script starts** and connects to Outlook
2. **Finds emails** with subject "EDF GDC PICKS - MM/DD/YYYY"
3. **Extracts date** from email subject (e.g., 12/15/2025)
4. **Parses HTML tables** from email body
5. **Identifies chamber types** (Dry, Frozen Dairy Deli, Meat & Produce)
6. **Maps columns** from email table to Excel columns
7. **Adds rows** to the "Daily Data" worksheet with:
   - Date (from email subject) in column A
   - All data columns from the email table
8. **Saves Excel** file to OneDrive
9. **Deletes emails** (in live mode only)

## Important Notes

### Dry Run Mode
- Always test first with dry run mode
- Shows exactly what would be done
- No changes are made to Excel or emails
- Start here to verify the process works

### Live Mode
- Makes actual changes
- Updates Excel workbook
- Deletes processed emails
- Requires confirmation before running

### Excel Workbook
- **File Location**: `C:\Users\tjschol\OneDrive - Walmart Inc\1. EDF\FY26 EDF Data_v.3.1.xlsx`
- **Worksheet**: "Daily Data"
- **Column A**: Date (automatically added from email subject)
- **Columns B+**: Data from email tables

### Email Requirements
- **Subject Format**: Must contain "EDF GDC PICKS" and a date (MM/DD/YYYY)
- **Body Format**: Must contain HTML tables
- **Tables**: One table per chamber (Dry, Frozen Dairy Deli, Meat & Produce)

## Troubleshooting

### "ModuleNotFoundError: No module named 'win32com'"
- The virtual environment needs to be reinitialized
- Run: `.venv\Scripts\pip install pywin32`

### "Excel file not found"
- Verify the OneDrive folder is synced locally
- Check the path: `C:\Users\tjschol\OneDrive - Walmart Inc\1. EDF\`

### "No tables found in email body"
- The email HTML structure might be different
- Check that the email contains HTML tables (not just plain text)

### "Chamber type not identified"
- The table context doesn't clearly indicate chamber type
- Verify email includes chamber name near the table

## Future Automation

### Option 1: Windows Task Scheduler

1. Open Task Scheduler
2. Create Basic Task
3. Set trigger (Daily at specific time)
4. Set action to run: `C:\Users\tjschol\Documents\edf-email-automation\run_live.bat`
5. Task will run automatically at scheduled time

### Option 2: Python Scheduler

Modify `main.py` to add scheduled execution:

```python
import schedule
import time

schedule.every().day.at("09:00").do(main, dry_run=False)

while True:
    schedule.run_pending()
    time.sleep(60)
```

Then run with: `.venv\Scripts\python.exe main_scheduler.py`

## Support

If you encounter issues:

1. First, run in **dry run mode** to see what would happen
2. Check the **console output** for specific error messages
3. Verify your **email subject** contains both "EDF GDC PICKS" and a date
4. Ensure **Outlook** is running before executing the script
5. Check **OneDrive** is synced and the Excel file is accessible

## Script Anatomy

### outlook_handler.py
- Connects to Outlook via COM
- Finds emails matching the pattern
- Extracts dates from subjects
- Deletes emails

### table_parser.py
- Parses HTML tables
- Identifies chamber types
- Extracts headers and data rows

### excel_handler.py
- Loads Excel workbook
- Maps email columns to Excel columns
- Adds rows with proper data types
- Saves workbook

### main.py
- Orchestrates the workflow
- Provides user feedback
- Handles dry run vs live mode
- Error handling and reporting

## Performance

- Processing 3 emails: ~10-15 seconds
- Excel save: ~5 seconds
- Total: ~15-20 seconds per run

## Version

- Version 1.0
- Created: December 2025
- Python 3.13.5