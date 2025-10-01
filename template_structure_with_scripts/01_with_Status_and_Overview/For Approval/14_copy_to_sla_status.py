import sys
import os
import shutil
from pathlib import Path

# Add parent directory (where config.py is located) to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..', '..'))  # since this script is in a subfolder's subfolder
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Import config.py from parent directory

# Use config variables instead of hardcoded ones
previous_month_num = config.previous_month_num
current_month_num = config.current_month_num
year = config.year

countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']

# Base path parts before the monthly scans folder
base_path_prefix = Path(r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning")

# Construct the dynamic monthly scans folder name, e.g. "25_08 Monthly Scans"
year_suffix = year[-2:]  # last two digits of year, e.g. "25"
monthly_scans_folder_name = f"{year_suffix}_{current_month_num} Monthly Scans"

# Full base destination path
base_dest_path = base_path_prefix / monthly_scans_folder_name / "01_with_Status_and_Overview" / "SLA Status"

# Current folder where the files to copy are located
current_folder = Path.cwd()  # Change this if your files are in a different folder

def get_destination_folder(country: str) -> Path:
    """
    Returns the destination folder path for a given country.
    """
    return base_dest_path / f"{country} Qualys Monthly Scans"

def copy_files():
    """
    Copies files from the current folder to their respective country folders based on filename pattern.
    """
    for file_path in current_folder.iterdir():
        if file_path.is_file():
            filename = file_path.name
            # Check if the filename contains a country and the year_month pattern
            for country in countries:
                if country in filename and f"{year}_{current_month_num}" in filename:
                    dest_folder = get_destination_folder(country)
                    if not dest_folder.exists():
                        print(f"Destination folder does not exist, creating: {dest_folder}")
                        dest_folder.mkdir(parents=True, exist_ok=True)
                        print(f"Created folder: {dest_folder}")

                    dest_file = dest_folder / filename
                    try:
                        shutil.copy2(file_path, dest_file)
                        print(f"Copied '{filename}' to '{dest_folder}'")
                    except Exception as e:
                        print(f"Error copying '{filename}': {e}")
                    break  # Stop checking other countries once matched

if __name__ == "__main__":
    print(f"Starting file copy for Year: {year}, Month: {current_month_num}")
    print(f"Source folder: {current_folder}")
    print(f"Destination base path: {base_dest_path}")
    copy_files()
    print("File copy completed.")
