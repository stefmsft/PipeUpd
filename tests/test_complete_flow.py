#!/usr/bin/env python3
"""
Test script for complete data flow: df_master -> shift -> df_pipe -> Excel
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import UpdatePipe
import pandas as pd

def test_complete_data_flow():
    """Test the complete data flow including the mapping to df_pipe"""
    print('Testing complete data flow: df_master -> shift -> df_pipe...')

    # Step 1: Create history DataFrame
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

    # Step 2: Create master DataFrame (simulates original Excel data)
    df_master = pd.DataFrame({
        'Key': ['TEST1'],
        'Week 37': ['tata'],  # Original data
        'Week 38': ['titi'],
        'Week 39': ['toto'],
        'Week 40': ['tutu'],
        'Week 41': ['tete']
    })

    print('Original df_master:')
    existing_cols = ['Week 37', 'Week 38', 'Week 39', 'Week 40', 'Week 41']
    print(df_master.iloc[0][existing_cols].to_dict())

    # Step 3: Apply shift to df_master (simulates the shift operation)
    new_week_columns = ['Week 39', 'Week 40', 'Week 41', 'Week 42', 'Week 43']
    df_master_shifted = UpdatePipe.ApplyWeekShiftFromHistory(df_master, df_whisto, new_week_columns)

    print('\\nShifted df_master:')
    print(df_master_shifted.iloc[0][existing_cols].to_dict())

    # Step 4: Create df_pipe and map data from shifted df_master (simulates the mapping step)
    df_pipe = pd.DataFrame({
        'Key': ['TEST1'],
        'Opportunity Number': ['OPT123'],
        'Sales Model Name': ['ModelABC']
    })

    # Temporarily set the global df_master for Mapping_Generic to work
    UpdatePipe.df_master = df_master_shifted

    # Apply the week column mapping logic (this is what happens in UpdatePipe.py lines 1296-1308)
    dynamic_week_columns = new_week_columns
    existing_week_columns = ['Week 37', 'Week 38', 'Week 39', 'Week 40', 'Week 41']

    for i, new_week_col in enumerate(dynamic_week_columns):
        if i < len(existing_week_columns):
            old_week_col = existing_week_columns[i]  # Positional column
            # Use the new fixed logic
            df_pipe[new_week_col] = df_pipe['Key'].apply(
                lambda key: UpdatePipe.Mapping_Generic(key, old_week_col)
            )

    print('\\nFinal df_pipe (what goes to Excel):')
    pipe_result = {}
    for new_col in new_week_columns:
        if new_col in df_pipe.columns:
            pipe_result[new_col] = df_pipe.iloc[0][new_col]
    print(pipe_result)

    # Step 5: Verify the expected results
    expected = {
        'Week 39': 'toto',  # Should get Week 39 data from history
        'Week 40': 'tutu',  # Should get Week 40 data from history
        'Week 41': 'tete',  # Should get Week 41 data from history
        'Week 42': '',      # Should be empty (no history)
        'Week 43': ''       # Should be empty (no history)
    }

    print('\\nExpected:')
    print(expected)

    # Check results
    success = True
    for col, expected_val in expected.items():
        if col in pipe_result:
            actual_val = str(pipe_result[col]).strip()
            if actual_val != expected_val:
                print(f"\\nFAILED: Column {col} expected '{expected_val}', got '{actual_val}'")
                success = False
        else:
            print(f"\\nFAILED: Column {col} not found in result")
            success = False

    if success:
        print('\\nComplete data flow test PASSED!')
        print('Excel column V (Week 39) should now contain: toto')
    else:
        print('\\nComplete data flow test FAILED!')

    return success

if __name__ == "__main__":
    try:
        if test_complete_data_flow():
            print("\\nTest PASSED - Excel output should be correct now!")
        else:
            print("\\nTest FAILED - Issue still exists")
    except Exception as e:
        print(f"Test failed with error: {str(e)}")
        import traceback
        traceback.print_exc()