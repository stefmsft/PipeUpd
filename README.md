# PipeUpdUV

A Python tool for integrating Salesforce pipeline exports into Excel tracking files while preserving manually entered data.

## Prerequisites

- Windows PowerShell (for setup scripts)
- Internet connection (for dependency installation)

## Quick Start

### 1. Initial Setup

```powershell
.\Setup.ps1
```

This will:
- Install Chocolatey, Git, and Python 3
- Create Python virtual environment 
- Install required dependencies
- Create .env configuration file from template

### 2. Configuration

Edit the `.env` file with your file paths:

```env
DIRECTORY_PIPE_RAW=C:\path\to\salesforce\exports
INPUT_SUIVI_RAW=C:\path\to\input\tracking.xlsm
OUTPUT_SUIVI_RAW=C:\path\to\output\tracking.xlsm
```

### 3. Usage

**Interactive Mode:**
```powershell
.\Run.ps1
```

**Direct Execution:**
```bash
# Process latest Salesforce export
python UpdatePipe.py

# Process specific file  
python UpdatePipe.py "C:\path\to\specific\export.xlsx"

# Process all files in directory
python UpdatePipe.py all
```

## What It Does

- Merges Salesforce opportunity data into Excel tracking spreadsheet
- Preserves existing manual data in key columns:
  - Estimated quantities
  - Revenue projections  
  - Invoice quarters
  - Forecast confidence levels
  - Support comments
- Filters and cleanses data automatically
- Updates pipeline analysis with trending charts
- Maintains historical logs for tracking

## Features

- **Data Preservation**: Manual entries are never overwritten
- **Smart Filtering**: Removes test data, invalid entries, and excluded owners
- **Automated Analysis**: Rolling window analysis with configurable timeframes
- **Backup Support**: Optional file backup before processing
- **Error Handling**: Comprehensive logging and error reporting

## Configuration Options

Key `.env` settings:

| Setting | Purpose | Default |
|---------|---------|---------|
| `SKIP_ROW` | Header rows to skip in Salesforce exports | 12 |
| `ROLLINGWINDOWS` | Analysis window size | 31 |
| `BCKUP_PIPE_FILE` | Enable backup before processing | False |

## Output

The tool updates the Excel file with:
- **Pipeline Sell Out** sheet: Main opportunity data
- **Pipeline Run Rate** sheet: Run rate opportunities  
- **Pipe Log** sheet: Historical tracking data
- **Pipe Analysis** sheet: Trend analysis and charts
