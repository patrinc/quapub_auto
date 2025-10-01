import sys
import os
import shutil

# Get the current directory (where this script is located)
current_dir = os.path.dirname(os.path.abspath(__file__))

# Get the parent directory (where config.py is located)
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))

# Add the parent directory to sys.path so you can import config
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config

countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
report_types = ['Network', 'Server']
country_folders = {
    "Brazil": "Brazil Qualys Monthly Scans",
    "CSA": "CSA Qualys Monthly Scans",
    "India": "India Qualys Monthly Scans",
    "Mexico": "Mexico Qualys Monthly Scans",
    "SFTL": "SFTL Qualys Monthly Scans"
}

src_base = r"C:\Users\example.user1\OneDrive - exampledomain\Qualys Monthly Scans"
dst_folder = fr"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning\25_{config.current_month_num} Monthly Scans\00_Previous_Final_Reports_And_Exceptions"

os.makedirs(dst_folder, exist_ok=True)

for country in countries:
    src_folder = os.path.join(src_base, country_folders[country])
    for report_type in report_types:
        filename = f"{country}_{config.year}_{config.previous_month_num}_Final_Report_{report_type}_Report.xlsx"
        src_file_path = os.path.join(src_folder, filename)
        dst_file_path = os.path.join(dst_folder, filename)
        if os.path.isfile(src_file_path):
            try:
                shutil.copy2(src_file_path, dst_file_path)
                print(f"Copied: {src_file_path} -> {dst_file_path}")
            except Exception as e:
                print(f"Error copying {src_file_path}: {e}")
        else:
            print(f"File not found: {src_file_path}")
