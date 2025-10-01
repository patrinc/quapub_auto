import pandas as pd
import os
import warnings
from config import year, current_month_num

warnings.simplefilter(action='ignore', category=FutureWarning)

countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
report_types = ['Network', 'Server']

base_path = r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"
key_cols = ['IP', 'QID', 'Port']

def safe_read_excel(filepath):
    try:
        df = pd.read_excel(filepath)
        print(f"Loaded file: {os.path.basename(filepath)} | Rows: {len(df)}")
        return df
    except FileNotFoundError:
        print(f"Warning: File not found - {os.path.basename(filepath)}")
        return pd.DataFrame()
    except Exception as e:
        print(f"Error reading {os.path.basename(filepath)}: {e}")
        return pd.DataFrame()

def clean_port(val):
    if pd.isna(val) or str(val).strip().lower() == 'nan' or str(val).strip() == '':
        return ''
    try:
        num = float(val)
        if num.is_integer():
            num = int(num)
        return str(num)
    except:
        return str(val).strip()

def normalize_key_columns(df, cols):
    for col in cols:
        if col in df.columns:
            if col == 'Port':
                df[col] = df[col].apply(clean_port)
            elif col == 'QID':
                def clean_qid(val):
                    if pd.isna(val) or str(val).strip() == '':
                        return ''
                    try:
                        num = float(val)
                        if num.is_integer():
                            return str(int(num))
                        else:
                            return str(val).strip()
                    except:
                        return str(val).strip()
                df[col] = df[col].apply(clean_qid)
            else:
                df[col] = df[col].astype(str).str.strip().str.lower()
        else:
            print(f"Warning: Key column '{col}' missing in dataframe.")
    return df

def process_reports(country, report_type):
    year_short = year[-2:]
    curr_raw_folder = f"{year_short}_{current_month_num} Monthly Scans"
    curr_raw_path = os.path.join(base_path, curr_raw_folder, f"{country}_{year}_{current_month_num}_Raw_Report_{report_type}_Report.xlsx")

    curr_raw = safe_read_excel(curr_raw_path)

    if curr_raw.empty:
        print(f"[{country} - {report_type}] Skipping due to missing input file.\n")
        return

    # Normalize key columns
    curr_raw = normalize_key_columns(curr_raw, key_cols)

    # Filter out rows with empty Solution
    curr_raw_filtered = curr_raw[~(curr_raw['Solution'].isna() | (curr_raw['Solution'].astype(str).str.strip() == ''))].copy()
    removed_empty_solution = len(curr_raw) - len(curr_raw_filtered)
    print(f"[{country} - {report_type}] Removed {removed_empty_solution} rows with empty 'Solution'.")

    # Remove rows with Type 'Ig' or 'Practice'
    before_type_filter = len(curr_raw_filtered)
    curr_raw_filtered = curr_raw_filtered[~curr_raw_filtered['Type'].isin(['Ig', 'Practice'])]
    removed_type = before_type_filter - len(curr_raw_filtered)
    print(f"[{country} - {report_type}] Removed {removed_type} rows with Type 'Ig' or 'Practice'.")

    # Remove rows with Severity == 1
    before_severity_filter = len(curr_raw_filtered)
    curr_raw_filtered = curr_raw_filtered[curr_raw_filtered['Severity'] != 1]
    removed_severity = before_severity_filter - len(curr_raw_filtered)
    print(f"[{country} - {report_type}] Removed {removed_severity} rows with Severity 1.")

    # Mark remaining rows as 'Reviewed'
    curr_raw_filtered['Status'] = 'Reviewed'

    # Ensure 'Comments' column exists
    if 'Comments' not in curr_raw_filtered.columns:
        insert_pos = curr_raw_filtered.columns.get_loc('Status') + 1
        curr_raw_filtered.insert(insert_pos, 'Comments', '')

    # Sort by key columns for consistency
    final_sorted = curr_raw_filtered.sort_values(by=key_cols)

    # Save final report
    output_folder = "01_with_Status_and_Overview"
    os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, f"{country}_{year}_{current_month_num}_Final_Report_with_Status_{report_type}_Report.xlsx")

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        final_sorted.to_excel(writer, index=False, sheet_name='Final Report')

    print(f"[{country} - {report_type}] Final report saved: {output_file}\n")

def main():
    for country in countries:
        for report_type in report_types:
            print(f"Processing {country} - {report_type} reports...")
            process_reports(country, report_type)
    print("All reports processed successfully.")

if __name__ == "__main__":
    main()
