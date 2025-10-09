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
- **Pipeline Close Lost** sheet: Closed lost opportunities
- **Week History** sheet: Complete historical tracking of all week data (W01-W53)
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
