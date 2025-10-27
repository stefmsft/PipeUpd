# PipeUpdUV

A Python tool for integrating Salesforce pipeline exports into Excel tracking files while preserving manually entered data.

## Prerequisites

- Windows PowerShell (for setup scripts)
- Internet connection (for dependency installation)

### PowerShell Version Recommendation

**For best Unicode support and modern features**, we recommend upgrading to PowerShell 7.5+:

**Easiest Installation Method (via Winget):**
```powershell
# Open PowerShell as Administrator and run:
winget install --id Microsoft.PowerShell --source winget
```

**Alternative Installation Methods:**

1. **Microsoft Store** (Simplest for Windows 10/11):
   - Open Microsoft Store
   - Search for "PowerShell"
   - Click "Install"

2. **Manual Installer**:
   - Download from: https://github.com/PowerShell/PowerShell/releases/latest
   - Choose the `.msi` installer for your system (x64 or ARM64)
   - Run the installer with default settings

**After Installation:**
- PowerShell 7.5+ installs alongside Windows PowerShell (doesn't replace it)
- Launch via Start Menu → "PowerShell 7" or run `pwsh` command
- Enable Unicode icons by adding `ENABLE_UNICODE=true` to your `.env` file

**Why Upgrade?**
- ✅ Full Unicode support (emojis display correctly)
- ✅ Better performance and modern features
- ✅ Cross-platform compatibility
- ✅ Long-term support and updates

**Note:** If you prefer to stay on Windows PowerShell 5.x, the scripts will work fine with ASCII fallback icons (`[!]` instead of `⚠️`).

## Important: Git Clone Setup if not running Powershell 7.5+

**If you clone this repository on a new device** and you are not using Powershell 7.5+, you MUST run these commands to ensure proper file encoding:

```powershell
# 1. Navigate to the project directory
cd C:\Projects\PipeUpd

# 2. Reset git attributes to fix encoding
git rm --cached -r .
git reset --hard HEAD

# 3. Verify PowerShell script encoding (should show UTF-8 with BOM)
Get-Content ProjectSetup.ps1 -Encoding UTF8 | Select-Object -First 1
```

**Why?** The PowerShell scripts contain Unicode emoji characters (🚀, ✅, etc.) that require UTF-8 encoding with BOM. The `.gitattributes` file ensures correct encoding, but git needs to re-apply it after cloning.

**Symptoms of encoding issues:**
- Parse errors mentioning "Jeton inattendu" or "unexpected token"
- Garbled characters like "âœ…" in error messages
- Scripts fail immediately when run

## Quick Start

### 1. Initial Setup

```powershell
.\ProjectSetup.ps1
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

- **Automatic Header Detection**: Intelligently detects Salesforce export header rows (no manual configuration needed!)
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

- **Automatic Header Detection**: 🆕 Intelligently scans Salesforce exports to find header rows
- **Data Preservation**: Manual entries are never overwritten
- **Smart Filtering**: Removes test data, invalid entries, and excluded owners
- **Automated Analysis**: Rolling window analysis with configurable timeframes
- **Backup Support**: Optional file backup before processing
- **Error Handling**: Comprehensive logging and error reporting

### Automatic Header Detection 🆕

The system now automatically detects where the actual data headers start in Salesforce export files, eliminating the need for manual `SKIP_ROW` configuration.

**How it works:**
- Scans the first 30 rows of the Excel file
- Looks for key header columns: "Opportunity Owner", "Created Date", "Close Date"
- Automatically determines how many rows to skip
- Adapts to Salesforce export format changes (warning lines, etc.)

**Benefits:**
- ✅ No manual configuration needed
- ✅ Works with varying Salesforce export formats
- ✅ Handles warning lines automatically (e.g., "Exported first 15,000 rows...")
- ✅ Backward compatible with existing SKIP_ROW setting

**Testing:**
```bash
# Run header detection tests
uv run python tests/test_header_detection.py
```

The test validates detection with both standard exports (12 header rows) and exports with warning lines (15+ header rows).

### Colored Logging 🎨

The application uses color-coded logging for better visibility and quick issue identification:

**Log Level Colors:**
- **INFO**: Default text (white) - Normal operation messages
- **DEBUG**: Yellow 🟡 - Detailed debugging information
- **WARNING**: Cyan 🔵 - Important notices and deprecation warnings
- **ERROR**: Red 🔴 - Error conditions requiring attention

**Key Log Messages:**
- Source directory path (cyan)
- Pipe file name (green)
- Auto-detected header row (info)
- SKIP_ROW deprecation notices (cyan warning)

**Example Output:**
```
INFO - Source directory: C:\Projects\PipeUpdUV\tests
INFO - Using pipe file: ASUS BTB PIPELINE - Stef-2025-10-27.xlsx
INFO - Auto-detected header row at line 16 (will skip 15 rows)
WARNING - SKIP_ROW is deprecated. Auto-detection is now used by default.
```

**Enable Debug Logging:**
Add to your `.env` file:
```env
LOG_LEVEL=DEBUG
```

## Configuration Options

Key `.env` settings:

| Setting | Purpose | Default |
|---------|---------|---------|
| ~~`SKIP_ROW`~~ | **[DEPRECATED]** Header rows to skip (now auto-detected) | Auto |
| `ROLLINGWINDOWS` | Analysis window size | 31 |
| `BCKUP_PIPE_FILE` | Enable backup before processing | False |

**Note:** `SKIP_ROW` is deprecated as of V2.0. The system now uses automatic header detection. If specified, it will be used as a fallback if auto-detection fails.

## Output

The tool updates the Excel file with:
- **Pipeline Sell Out** sheet: Main opportunity data
- **Pipeline Run Rate** sheet: Run rate opportunities
- **Pipeline Close Lost** sheet: Closed lost opportunities
- **Week History** sheet: Complete historical tracking of all week data (W01-W53)
- **Owner Opty Tracking** sheet: Unique opportunity counts per owner per week (W01-W53)
- **Pipe Log** sheet: Historical tracking data
- **Pipe Analysis** sheet: Trend analysis and charts

## Week History & Dynamic Shifting

The system now includes advanced week management features:

### Week History Tracking
- **Complete Archive**: All week data is preserved in the "Week History" tab with columns W01-W53
- **Data Preservation**: Before any shifts occur, current week data is copied to the history
- **Key-Based Storage**: Each opportunity is tracked by its unique key (Opportunity Number + Sales Model Name)

### Dynamic Week Shifting
- **Auto-Detection**: System detects when the current week has changed from the center column (X)
- **Smart Shifting**: Data automatically shifts based on the new week range while preserving historical mappings
- **Data Integrity**: Uses Week History as the source of truth for accurate week-to-data mapping

### How It Works
1. **Detection**: Compare current week vs center column (Week X) to calculate shift amount
2. **Backup**: Copy all existing week data to Week History before any changes
3. **Shift**: Apply calculated shift using historical data for accurate mapping
4. **Update**: Refresh Excel headers and data with new week range

Example: If last run was centered on Week 39 and current week is 41:
- **Shift Amount**: +2 weeks
- **Old Range**: Week 37-41 → **New Range**: Week 39-43
- **Data Mapping**: Week 39 data (from history) → Column V (labeled "Week 39")

## Testing

The project includes comprehensive test suites to validate functionality:

### Test Files Location
```
tests/
├── test_week_shift.py      # Week shifting and history functionality
└── test_complete_flow.py   # End-to-end data flow validation
```

### Running Tests with uv

**Prerequisites**: Ensure uv is installed and the project dependencies are available.

**Execute Individual Tests:**
```bash
# Test week shifting functionality
uv run python tests/test_week_shift.py

# Test complete data flow (df_master -> df_pipe -> Excel)
uv run python tests/test_complete_flow.py
```

**Execute All Tests:**
```bash
# Run all test files
uv run python -m pytest tests/ -v

# Or run them individually
uv run python tests/test_week_shift.py && uv run python tests/test_complete_flow.py
```

### Test Coverage

**test_week_shift.py** validates:
- Week shift detection logic (current week vs center column)
- Week History DataFrame creation and management
- Historical data mapping functionality
- Week-to-data preservation during shifts

**test_complete_flow.py** validates:
- Complete data flow from df_master through df_pipe to Excel output
- Correct mapping of shifted week data to final output columns
- Integration between history-based shifting and Excel writing

### Test Development Guidelines

When developing new tests:
1. **Use uv for execution**: Always run tests with `uv run python` to ensure proper environment
2. **Add path context**: Tests include path setup to import UpdatePipe module
3. **Create realistic data**: Use test DataFrames that mirror actual Excel structure
4. **Validate end-to-end**: Test the complete flow, not just individual functions
5. **Include edge cases**: Test year boundaries, missing data, and shift scenarios

### Expected Test Output

Successful test runs show:
```
Testing DetectWeekShift function...
Week shift detection test passed

Testing Week History functions...
Week History functions test passed

Testing actual scenario with history: Week 37-41 -> Week 39-43...
Actual scenario with history test PASSED!

All tests PASSED!
```

### Debugging Tests

For verbose logging during tests:
```bash
# Enable debug logging
uv run python -c "
import logging
logging.getLogger().setLevel(logging.DEBUG)
exec(open('tests/test_week_shift.py').read())
"
```

The test suite ensures that:
- Week data maintains correct week-to-value relationships during shifts
- Historical data is preserved and accessible
- Excel output matches expected week mapping
- System handles various shift scenarios (forward, backward, year boundaries)

## Owner Opportunity Tracking

The system tracks unique opportunities created per week by each sales owner:

### Features
- **Unique Counting**: Counts distinct opportunity numbers (not total rows)
- **Weekly Granularity**: Tracks opportunities by ISO week number (W01-W53)
- **Year Filtering**: Only counts opportunities from the current year
- **Owner Filtering**: Supports excluding specific owners via configuration
- **Maximum Preservation**: Never decreases counts - always keeps the maximum value seen
- **Persistent Storage**: Owner rows are never deleted, even if owner has no current opportunities

### Configuration

Exclude specific owners from tracking in `.env`:
```env
# Comma-separated list of owners to exclude
EXCLUDED_OPTY_OWNERS=John DOE,Jane SMITH,Old Owner
```

### Debugging Tool

The `debug_owner_week.py` script helps investigate opportunity counts for specific owners and weeks:

**Usage:**
```bash
# Check opportunities for a specific owner and week
python debug_owner_week.py "Owner Name" 43

# Using uv
uv run python debug_owner_week.py "John DOE" 41
```

**What it shows:**
- Total opportunities found for the owner
- Unique opportunity count for the specified week
- Detailed list of each unique opportunity with:
  - Opportunity number
  - Customer name
  - Quantity and price
  - Creation date
- Warnings for future-dated opportunities
- Duplicate detection (same opportunity on multiple rows)

**Example output:**
```
================================================================================
Searching for opportunities: Owner='John DOE', Week=41
================================================================================

Loading pipe file: ASUS BTB PIPELINE - Stef-2025-10-10-06-00-11.xlsx
Found 981 total opportunities for 'John DOE'
Found 8 total rows in Week 41 of 2025
Found 2 duplicate opportunity numbers (keeping max values)
Unique opportunities: 6

Owner                     Opty Number     Customer                       Qty        Total Price     Created Date
--------------------------------------------------------------------------------------------------------------
John DOE                OP0000271712    Mairie de Fort de x            1          €1,719          2025-10-06
John DOE                OP0000271714    Mairie de Fort de x            1          €1,719          2025-10-06
...

Total unique opportunities: 6
```

**Common use cases:**
- Verify opportunity counts match between tab and source data
- Investigate discrepancies in weekly tracking
- Identify future-dated or duplicate opportunities
- Understand which opportunities are being counted for a specific owner/week
