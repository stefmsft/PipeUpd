"""
Debug script to dump opportunities for a specific owner and week

Usage:
    python debug_owner_week.py "Owner Name" 43

This will show all opportunities created by the owner in the specified week.
"""

import sys
import pandas as pd
import os
from datetime import datetime
from dotenv import load_dotenv
import platform

# Load environment variables
load_dotenv()

# Detect if we can use Unicode icons safely
# On Windows, check PowerShell version. Use plain text for older versions.
def can_use_unicode():
    """
    Check if Unicode icons can be displayed safely.

    Returns True if:
    - Not on Windows (Linux/Mac handle Unicode well)
    - ENABLE_UNICODE environment variable is set to 'true' or '1'

    Otherwise returns False on Windows (safe default for old PowerShell)
    """
    # Allow explicit override via environment variable
    enable_unicode = os.getenv('ENABLE_UNICODE', '').lower()
    if enable_unicode in ('true', '1', 'yes'):
        return True

    # Non-Windows systems typically handle Unicode well
    if platform.system() != 'Windows':
        return True

    # On Windows, default to False (safe for PowerShell 5.x)
    # Users with PowerShell 7.5+ can set ENABLE_UNICODE=true in .env
    return False

USE_UNICODE = can_use_unicode()
WARNING_ICON = "⚠️" if USE_UNICODE else "[!]"

DIRECTORY_PIPE_RAW = os.getenv("DIRECTORY_PIPE_RAW")
SKIP_ROW = int(os.getenv("SKIP_ROW", 12))

# Column indices (after reorg)
COL_OPTYOWNER = 0
COL_CREATED = 1
COL_CLOSED = 2
COL_STAGE = 3
COL_OPTYNUM = 4
COL_CUSTOMER = 6
COL_QTY = 7
COL_TOTPRICE = 9
COL_SALESMODELNAME = 10

def get_latest_pipe(directory):
    """Get the latest pipe file from directory"""
    import glob
    files = glob.glob(f'{directory}/*.xls*')
    if not files:
        print(f"ERROR: No Excel files found in {directory}")
        sys.exit(1)
    return max(files, key=os.path.getctime)

def main():
    if len(sys.argv) < 3:
        print("Usage: python debug_owner_week.py \"Owner Name\" week_number")
        print("Example: python debug_owner_week.py \"William ROMAN\" 43")
        sys.exit(1)

    target_owner = sys.argv[1]
    target_week = int(sys.argv[2])

    print(f"\n{'='*80}")
    print(f"Searching for opportunities: Owner='{target_owner}', Week={target_week}")
    print(f"{'='*80}\n")

    # Load the latest pipe file
    latest_pipe = get_latest_pipe(DIRECTORY_PIPE_RAW)
    print(f"Loading pipe file: {os.path.basename(latest_pipe)}")

    df_pipe = pd.read_excel(latest_pipe, skiprows=SKIP_ROW)

    # Drop unnamed columns
    unnamed_cols = [col for col in df_pipe.columns if str(col).startswith('Unnamed:')]
    if unnamed_cols:
        df_pipe.drop(columns=unnamed_cols, inplace=True)

    # Reorg columns (same as UpdatePipe.py)
    cols = list(df_pipe.columns.values)
    Cval = cols.pop(11)
    cols.insert(7, Cval)
    Cval = cols.pop(11)
    cols.insert(7, Cval)
    df_pipe = df_pipe.reindex(columns=cols)

    # Format dates
    df_pipe[cols[COL_CREATED]] = pd.to_datetime(df_pipe[cols[COL_CREATED]], format='mixed', errors='coerce')

    # Get column names
    owner_col = cols[COL_OPTYOWNER]
    created_col = cols[COL_CREATED]
    opty_col = cols[COL_OPTYNUM]
    customer_col = cols[COL_CUSTOMER]
    qty_col = cols[COL_QTY]
    price_col = cols[COL_TOTPRICE]

    # Filter for target owner
    owner_opties = df_pipe[df_pipe[owner_col] == target_owner].copy()

    if len(owner_opties) == 0:
        print(f"\nNo opportunities found for owner '{target_owner}'")
        print("\nAvailable owners in the pipe file:")
        unique_owners = df_pipe[owner_col].dropna().unique()
        for owner in sorted(unique_owners):
            print(f"  - {owner}")
        sys.exit(0)

    print(f"Found {len(owner_opties)} total opportunities for '{target_owner}'")

    # Get current year
    current_year = datetime.now().year

    # Add week and year columns
    owner_opties['Week'] = owner_opties[created_col].apply(
        lambda x: x.isocalendar()[1] if pd.notna(x) else None
    )
    owner_opties['Year'] = owner_opties[created_col].apply(
        lambda x: x.year if pd.notna(x) else None
    )

    # Filter for target week AND current year only
    week_opties = owner_opties[
        (owner_opties['Week'] == target_week) &
        (owner_opties['Year'] == current_year)
    ].copy()

    print(f"Found {len(week_opties)} total rows in Week {target_week} of {current_year}")

    if len(week_opties) == 0:
        print(f"No opportunities found for Week {target_week} in {current_year}")
        print(f"\nWeek distribution for '{target_owner}' in {current_year}:")
        current_year_opties = owner_opties[owner_opties['Year'] == current_year]
        week_counts = current_year_opties['Week'].value_counts().sort_index()
        for week, count in week_counts.items():
            if pd.notna(week):
                print(f"  Week {int(week):02d}: {count} opportunities")
        sys.exit(0)

    # Group by Opportunity Number and keep max values
    # First, ensure numeric columns are properly typed
    week_opties[qty_col] = pd.to_numeric(week_opties[qty_col], errors='coerce')
    week_opties[price_col] = pd.to_numeric(week_opties[price_col], errors='coerce')

    # Group by opportunity number and aggregate
    grouped = week_opties.groupby(opty_col).agg({
        owner_col: 'first',
        customer_col: 'first',
        qty_col: 'max',  # Keep maximum quantity
        price_col: 'max',  # Keep maximum price
        created_col: 'first'
    }).reset_index()

    unique_count = len(grouped)
    duplicate_count = len(week_opties) - unique_count

    if duplicate_count > 0:
        print(f"Found {duplicate_count} duplicate opportunity numbers (keeping max values)")
    print(f"Unique opportunities: {unique_count}\n")

    # Display results
    print(f"{'Owner':<25} {'Opty Number':<15} {'Customer':<30} {'Qty':<10} {'Total Price':<15} {'Created Date':<12}")
    print("-" * 110)

    today = datetime.now()
    future_count = 0

    for _, row in grouped.iterrows():
        owner = str(row[owner_col])[:24]
        opty = str(row[opty_col])[:14]
        customer = str(row[customer_col])[:29]
        qty = f"{row[qty_col]:.0f}" if pd.notna(row[qty_col]) else ""
        price = f"€{row[price_col]:,.0f}" if pd.notna(row[price_col]) else ""
        created = row[created_col].strftime('%Y-%m-%d') if pd.notna(row[created_col]) else ""

        # Check if future date
        is_future = ""
        if pd.notna(row[created_col]) and row[created_col] > today:
            is_future = f" {WARNING_ICON} FUTURE"
            future_count += 1

        print(f"{owner:<25} {opty:<15} {customer:<30} {qty:<10} {price:<15} {created:<12}{is_future}")

    print("-" * 110)
    print(f"\nTotal unique opportunities: {unique_count}")
    if future_count > 0:
        print(f"{WARNING_ICON} WARNING: {future_count} opportunities have FUTURE creation dates!")

    print(f"\n{'='*80}\n")

if __name__ == "__main__":
    main()
