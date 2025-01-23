import pandas as pd
import os
import sys
from pathlib import Path

# Get the absolute path to the script's directory
script_dir = Path(__file__).parent.absolute()

# Get the path to the Excel file
excel_file = os.path.join(script_dir, '..', 'tmp', 'r.xlsx')

# Read the Excel file
try:
    # Read all sheets into a dictionary of DataFrames
    excel_data = pd.read_excel(excel_file, sheet_name=None)
    
    # Print information about each sheet
    for sheet_name, df in excel_data.items():
        print(f"\nSheet: {sheet_name}")
        print("-" * 40)
        print("Shape:", df.shape)
        print("\nFirst few rows:")
        print(df.head())
        print("\n")

except FileNotFoundError:
    print(f"Error: Could not find Excel file at {excel_file}")
    print("Make sure to run the Dart example first to create the Excel file.")
except Exception as e:
    print(f"Error reading Excel file: {str(e)}")
