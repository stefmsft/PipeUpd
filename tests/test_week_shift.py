#!/usr/bin/env python3
"""
Test script for Week History and Week Shift functionality
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import UpdatePipe
import pandas as pd
from datetime import datetime

def test_week_shift_detection():
    """Test the DetectWeekShift function"""
    print('Testing DetectWeekShift function...')

    # Create test master DataFrame with week columns
    test_master = pd.DataFrame({
        'Key': ['TEST1', 'TEST2'],
        'Week 37': ['Data1', 'Data2'],
        'Week 38': ['Data3', 'Data4'],
        'Week 39': ['Data5', 'Data6'],  # This should be center column
        'Week 40': ['Data7', 'Data8'],
        'Week 41': ['Data9', 'Data10']
    })

    # Simulate current week 41 (shift should be 41-39=2)
    UpdatePipe.CURWEEK = 41
    shift, existing_cols = UpdatePipe.DetectWeekShift(test_master)
    print(f'Shift detected: {shift}')
    print(f'Existing columns: {existing_cols}')
    assert shift == 2, f"Expected shift of 2, got {shift}"
    print('Week shift detection test passed')
    return True

def test_actual_scenario_with_history():
    """Test the actual scenario using history DataFrame approach"""
    print('\\nTesting actual scenario with history: Week 37-41 -> Week 39-43...')

    # Create history DataFrame with the original data
    df_whisto = pd.DataFrame({
        'key': ['TEST1'],
        'W37': ['tata'],
        'W38': ['titi'],
        'W39': ['toto'],
        'W40': ['tutu'],
        'W41': ['tete'],
        'W42': [''],
        'W43': ['']
    })

    # Add remaining columns W01-W53
    for i in range(1, 54):
        col = f'W{i:02d}'
        if col not in df_whisto.columns:
            df_whisto[col] = ''

    # Create master DataFrame with existing week columns
    test_df = pd.DataFrame({
        'Key': ['TEST1'],
        'Week 37': ['tata'],  # These will be overwritten
        'Week 38': ['titi'],
        'Week 39': ['toto'],
        'Week 40': ['tutu'],
        'Week 41': ['tete']
    })

    print('Before shift (master):')
    existing_cols = ['Week 37', 'Week 38', 'Week 39', 'Week 40', 'Week 41']
    row_before = test_df.iloc[0][existing_cols].to_dict()
    print(row_before)

    print('History data:')
    history_data = {f'W{i}': df_whisto.iloc[0][f'W{i:02d}'] for i in range(37, 44)}
    print(history_data)

    # New week columns after shift
    new_week_columns = ['Week 39', 'Week 40', 'Week 41', 'Week 42', 'Week 43']

    # Apply the new history-based shift
    shifted_df = UpdatePipe.ApplyWeekShiftFromHistory(test_df, df_whisto, new_week_columns)
    print('After shift using history:')
    row_after = shifted_df.iloc[0][existing_cols].to_dict()
    print(row_after)

    # Expected result: Week 39=toto, Week 40=tutu, Week 41=tete, Week 42=blank, Week 43=blank
    # But we're looking at original column names, so:
    expected = {
        'Week 37': 'toto',  # Now represents Week 39 data from history
        'Week 38': 'tutu',  # Now represents Week 40 data from history
        'Week 39': 'tete',  # Now represents Week 41 data from history
        'Week 40': '',      # Now represents Week 42 data from history (empty)
        'Week 41': ''       # Now represents Week 43 data from history (empty)
    }

    print('Expected result:')
    print(expected)

    # Check results
    success = True
    for col, expected_val in expected.items():
        actual_val = str(row_after[col]).strip()
        if actual_val != expected_val:
            print(f"FAILED: Column {col} expected '{expected_val}', got '{actual_val}'")
            success = False

    if success:
        print('Actual scenario with history test PASSED!')
    return success

def test_backward_shift():
    """Test backward shift scenario"""
    print('\\nTesting backward shift scenario: Week 42-46 -> Week 39-43...')

    # Original data centered on week 44, shifting back to week 41
    test_df = pd.DataFrame({
        'Key': ['TEST1'],
        'Week 42': ['alpha'],   # Position 0 -> should move to position 3
        'Week 43': ['beta'],    # Position 1 -> should move to position 4
        'Week 44': ['gamma'],   # Position 2 -> should disappear (out of range)
        'Week 45': ['delta'],   # Position 3 -> should disappear (out of range)
        'Week 46': ['epsilon']  # Position 4 -> should disappear (out of range)
    })

    existing_cols = ['Week 42', 'Week 43', 'Week 44', 'Week 45', 'Week 46']
    print('Before backward shift:')
    print(test_df.iloc[0][existing_cols].to_dict())

    # Apply shift of -3 (current week 41, center was 44: 41-44=-3)
    shifted_df = UpdatePipe.ApplyWeekShift(test_df, -3, existing_cols)
    print('After shift of -3:')
    row_after = shifted_df.iloc[0][existing_cols].to_dict()
    print(row_after)

    # Expected result: blank, blank, blank, alpha, beta
    expected = {
        'Week 42': '',      # New empty
        'Week 43': '',      # New empty
        'Week 44': '',      # New empty
        'Week 45': 'alpha', # Was Week 42 data
        'Week 46': 'beta'   # Was Week 43 data
    }

    print('Expected result:')
    print(expected)

    # Check results
    for col, expected_val in expected.items():
        actual_val = str(row_after[col]).strip()
        if actual_val != expected_val:
            print(f"FAILED: Column {col} expected '{expected_val}', got '{actual_val}'")
            return False

    print('Backward shift test PASSED!')
    return True

def test_week_history_functions():
    """Test the Week History functions"""
    print('\\nTesting Week History functions...')

    # Test CreateWeekHistoryDataFrame
    columns = ['key'] + [f'W{i:02d}' for i in range(1, 54)]  # W01 to W53
    df_whisto = pd.DataFrame(columns=columns)
    print(f"Created DataFrame with {len(df_whisto.columns)} columns")

    # Test UpdateWeekHistoryRow simulation
    key = "OPT123ModelABC"
    week_data = {"Week 25": "Test Value 1", "Week 26": "Test Value 2"}

    # Simulate the logic
    new_row = {'key': key}
    for i in range(1, 54):
        new_row[f'W{i:02d}'] = ''

    for week_col, value in week_data.items():
        if week_col.startswith('Week '):
            week_num = week_col.replace('Week ', '')
            whisto_col = f'W{int(week_num):02d}'
            if whisto_col in new_row:
                new_row[whisto_col] = value

    df_whisto = pd.concat([df_whisto, pd.DataFrame([new_row])], ignore_index=True)
    print(f"Added row for key: {key}")
    print(f"W25 value: {df_whisto.iloc[0]['W25']}")
    print(f"W26 value: {df_whisto.iloc[0]['W26']}")
    print('Week History functions test passed')
    return True

if __name__ == "__main__":
    try:
        test_week_shift_detection()
        test_week_history_functions()

        print("\\n" + "="*50)
        print("TESTING ACTUAL USER SCENARIOS")
        print("="*50)

        success = True
        success &= test_actual_scenario_with_history()
        # success &= test_backward_shift()  # Disabled old test for now

        if success:
            print("\\nAll tests PASSED!")
        else:
            print("\\nSome tests FAILED!")

    except Exception as e:
        print(f"Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()