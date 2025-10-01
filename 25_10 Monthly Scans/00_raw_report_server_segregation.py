import os
import pandas as pd
import csv
from openpyxl import Workbook
from openpyxl.styles import Alignment
from config import year, current_month_num

# Input raw data directory (monthly folder from config)
raw_base_dir = r"C:\Users\example.user1\OneDrive - exampledomain\General - GSRC Global\CIC\Qualys\Qualys Historic Raw Data"
raw_month_dir = f"{year}_{current_month_num}_Raw_Reports"
base_dir = os.path.join(raw_base_dir, raw_month_dir)

# Metadata directory (configured path with year/month from config)
year_short = year[2:]  # "25" from "2025"
metadata_base_dir = r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"
metadata_month_dir = f"{year_short}_{current_month_num} Monthly Scans"
metadata_dir_full = os.path.join(metadata_base_dir, metadata_month_dir,
                                 "01_with_Status_and_Overview", "All Metadata")
os.makedirs(metadata_dir_full, exist_ok=True)

# Current working directory (script running directory) for saving server reports
current_dir = os.getcwd()

def save_metadata_as_excel(file_path, metadata_dir, bu):
    metadata_lines = []
    with open(file_path, 'r', encoding='utf-8', newline='') as f:
        reader = csv.reader(f)
        for _ in range(6):
            try:
                row_values = next(reader)
                row_values = [val.strip().strip('"') for val in row_values]
                metadata_lines.append(row_values)
            except StopIteration:
                break

    wb = Workbook()
    ws = wb.active
    ws.title = 'Metadata'

    cells_to_convert = {
        (2, 7),  # G2
        (6, 3),  # C6
        (6, 4),  # D6
    }
    cells_to_wrap = {
        (6, 7),  # G6
    }

    for row_idx, row_values in enumerate(metadata_lines, start=1):
        for col_idx, value in enumerate(row_values, start=1):
            if (row_idx, col_idx) in cells_to_convert:
                try:
                    if '.' in value:
                        cell_value = float(value)
                    else:
                        cell_value = int(value)
                except ValueError:
                    cell_value = value
            else:
                cell_value = value
            cell = ws.cell(row=row_idx, column=col_idx, value=cell_value)
            if (row_idx, col_idx) in cells_to_wrap:
                cell.alignment = Alignment(wrap_text=True)

    ws.row_dimensions[6].height = 14.4
    metadata_file_name = f'{bu}_Server_Metadata.xlsx'
    metadata_file_path = os.path.join(metadata_dir, metadata_file_name)
    wb.save(metadata_file_path)
    print(f"Saved metadata Excel: {metadata_file_name}")

# Process CSV files directly inside the monthly raw data folder
for file_name in os.listdir(base_dir):
    file_path = os.path.join(base_dir, file_name)
    if os.path.isfile(file_path):
        if '_Metadata' in file_name:
            continue
        if file_path.endswith('.csv') and 'Server_Agents_Monthly_Scan' in file_name:
            print(f"Processing file: {file_name}")
            bu = file_name.split('_')[2]  # BU extraction as in original script
            save_metadata_as_excel(file_path, metadata_dir_full, bu)

            # Read first 6 lines to get value at D6 (row 6, col 4)
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = [next(f) for _ in range(6)]
            value_at_D6 = lines[5].split(',')[3].replace('"', '')
            skip_rows = int(value_at_D6) + 13

            # Read data skipping calculated rows
            df = pd.read_csv(file_path, skiprows=skip_rows, low_memory=False)

            excel_file_name = f'{bu}_{year}_{current_month_num}_Raw_Report_Server_Report.xlsx'
            excel_file_path = os.path.join(current_dir, excel_file_name)
            df.to_excel(excel_file_path, index=False)
            print(f"Saved data Excel: {excel_file_name}")

print('All server reports and metadata have been processed.')
