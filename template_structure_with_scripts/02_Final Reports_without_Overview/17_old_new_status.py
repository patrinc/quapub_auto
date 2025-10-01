import os
import sys
import pandas as pd
from pathlib import Path

# Add parent directory (where config.py is located) to sys.path
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Import config.py from parent directory

# Use config variables instead of hardcoded values
previous_month_num = config.previous_month_num
current_month_num = config.current_month_num
year = config.year
countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
report_types = ['Network', 'Server']

previous_month_dir = Path(
    f"C:/Users/example.user1/OneDrive - exampledomain/Assignments/VAPT/Infra Scanning/25_{current_month_num} Monthly Scans/00_Previous_Final_Reports_And_Exceptions")
current_month_dir = Path(
    f"C:/Users/example.user1/OneDrive - exampledomain/Assignments/VAPT/Infra Scanning/25_{current_month_num} Monthly Scans/02_Final Reports_without_Overview")


def get_ip_qid_set(filepath):
    df = pd.read_excel(filepath, usecols=["IP", "QID"])
    df['IP_QID'] = df['IP'].astype(str) + "_" + df['QID'].astype(str)
    return set(df['IP_QID'])


# Load previous month vulnerabilities
previous_month_vulns = {}
for country in countries:
    for report_type in report_types:
        filename = f"{country}_{year}_{previous_month_num}_Final_Report_{report_type}_Report.xlsx"
        filepath = previous_month_dir / filename
        if filepath.exists():
            previous_month_vulns[(country, report_type)] = get_ip_qid_set(filepath)
        else:
            print(f"Previous month file not found: {filepath}")

# Update only Status column in current month files
for country in countries:
    for report_type in report_types:
        filename = f"{country}_{year}_{current_month_num}_Final_Report_{report_type}_Report.xlsx"
        filepath = current_month_dir / filename
        if filepath.exists():
            df_curr = pd.read_excel(filepath)  # Read all columns and data
            prev_vulns = previous_month_vulns.get((country, report_type), set())
            # Create the IP_QID combination column temporarily
            df_curr['IP_QID'] = df_curr['IP'].astype(str) + "_" + df_curr['QID'].astype(str)
            # Determine new status value
            df_curr['Status'] = df_curr['IP_QID'].apply(
                lambda x: 'Old' if x in prev_vulns else 'New'
            )
            df_curr.drop(columns=['IP_QID'], inplace=True)  # Remove helper column
            # Save back with all original columns, only Status updated
            df_curr.to_excel(filepath, index=False)
            print(f"Updated 'Status' column in file: {filepath}")
        else:
            print(f"Current month file not found: {filepath}")
