"""
Test script: Generate a data pack from the sample raw file
and compare against the hand-built reference file.
"""
import os
import sys
from collections import OrderedDict

# Add parent dir to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from engine.generator import generate_data_pack

# Hardcoded mapping for the test data
config = {
    "raw_data_sheet": "Raw ARR Data",
    "raw_data_first_row": 2,      # First data row (row 1 = headers)
    "raw_data_last_row": 23498,   # Last data row
    "customer_id_col": "B",       # "Accounts Account ID"
    "date_col": "A",              # "Period"
    "arr_col": "F",               # "Revenue ARR"
    "attributes": OrderedDict([
        ("Customer Type", "C"),    # "Industry"
        ("Customer Size", "D"),    # "Accounts Segment"
        ("Country", "E"),          # "Country"
    ]),
    "time_granularity": "monthly",
    "fiscal_year_end_month": 12,
    "scale_factor": 1000,
    "filter_breakouts": [
        {
            "title": "Software & Services",
            "filters": {"Customer Type": "Software & Services"}
        }
    ],
}

# File paths
input_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                          '..', 'Data Pack AI Clean Version-Raw Only.xlsx')
output_file = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                           '..', 'Data Pack AI Generated Output.xlsx')

if __name__ == '__main__':
    print(f"Input:  {os.path.abspath(input_file)}")
    print(f"Output: {os.path.abspath(output_file)}")
    print()
    generate_data_pack(config, input_file, output_file)
    print(f"\nOutput saved to: {os.path.abspath(output_file)}")
