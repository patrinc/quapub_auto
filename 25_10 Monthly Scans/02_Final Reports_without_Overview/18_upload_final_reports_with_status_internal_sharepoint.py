import os
import sys
import shutil
from pathlib import Path

# Add parent directory (where config.py is located) to sys.path to allow import
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Now config variables are accessible

# Use config variables
previous_month = config.previous_month
current_month = config.current_month
previous_month_num = config.previous_month_num
current_month_num = config.current_month_num
year = config.year

# Input directory is current working directory
input_directory = Path.cwd()

# Base output directory (hardcoded)
base_output_dir = Path(r"C:\Users\example.user1\OneDrive - exampledomain\General - GSRC Global\CIC\Vulnerability Mgmt\Qualys Vulnerability Status Records")

# Subfolder name based on year and month variables from config
subfolder_name = f"{year[-2:]}_{current_month_num} Monthly Scans"

# Full output directory path
output_directory = base_output_dir / subfolder_name

# Create output directory if it doesn't exist
output_directory.mkdir(parents=True, exist_ok=True)

# --- Main execution loop ---
print(f"Scanning for files in: {input_directory}")
print(f"Copying files to: {output_directory}\n")
files_found = 0
for file_path in input_directory.glob("*Final_Report*.xlsx"):
    files_found += 1
    target_path = output_directory / file_path.name
    try:
        shutil.copy2(file_path, target_path)  # copy2 to preserve metadata
        print(f"Copied file: {file_path.name} to {target_path}")
    except Exception as e:
        print(f"ERROR copying file {file_path.name}: {e}")
if files_found == 0:
    print("No 'Final_Report.xlsx' files found in the current directory to copy.")
else:
    print(f"\nCopy complete. {files_found} files were copied.")
