import os
import sys
import pandas as pd

# Add parent directory (where config.py is) to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__)) 
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Import config.py from parent directory

# Variables from config
previous_month_num = config.previous_month_num
current_month_num = config.current_month_num
year = config.year

countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
report_types = ['Network', 'Server']

# Folder containing final reports
base_folder = "."

# Exception list folder path (remains hardcoded as before)
exception_list_folder = r"C:\Users\example.user1\OneDrive - exampledomain\General - GSRC Global\CIC\Vulnerability Mgmt\Qualys Global Exception Lists"

# File name pattern for reports
def final_report_filename(country, year, month, report_type):
    return f"{country}_{year}_{month}_Final_Report_with_Status_{report_type}_Report.xlsx"

# Exception list file path - always previous month
exception_list_file = os.path.join(exception_list_folder, f"Global_Exception_List_{previous_month_num}.xlsx")

# Load exception sheets
exception_sheets = pd.read_excel(exception_list_file, sheet_name=["Cleanup List", "Exception List"])

# Combine sheets into one DataFrame
exception_df = pd.concat([exception_sheets["Cleanup List"], exception_sheets["Exception List"]], ignore_index=True)

# Columns to use as keys for matching exceptions
key_columns = ["IP", "QID", "Title"]

# Create set of unique exception keys for fast lookup
exception_keys = set()
for _, row in exception_df.iterrows():
    key = tuple(row[col] for col in key_columns)
    exception_keys.add(key)

print(f"Loaded {len(exception_keys)} exception entries from previous month ({previous_month_num}) to filter.")

removal_counts = {}

for country in countries:
    for report_type in report_types:
        report_file = final_report_filename(country, year, current_month_num, report_type)
        report_path = os.path.join(base_folder, report_file)
        if not os.path.exists(report_path):
            print(f"File not found: {report_path}")
            continue
        try:
            df = pd.read_excel(report_path)
            original_count = len(df)

            def is_exception(row):
                key = tuple(row.get(col, None) for col in key_columns)
                return key in exception_keys

            filtered_df = df[~df.apply(is_exception, axis=1)]
            removed = original_count - len(filtered_df)
            removal_counts[report_file] = removed
            print(f"{report_file}: Removed {removed} rows as exceptions.")
            filtered_df.to_excel(report_path, index=False)
        except Exception as e:
            print(f"Error processing {report_file}: {e}")

print("\nSummary of removals per file:")
for file, count in removal_counts.items():
    print(f"{file}: {count} rows removed")
