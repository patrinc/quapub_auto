import os
import pandas as pd

# Define the current directory
directory = "."

# Function to process each Excel file
def process_excel_file(filepath):
    try:
        excel_file = pd.ExcelFile(filepath)
        
        # Exclude the sheets "Overview" and "Status Description"
        sheets_to_process = [s for s in excel_file.sheet_names if s not in ['Overview', 'Status Description']]
        
        processed_sheets = {}

        for sheet_name in sheets_to_process:
            sheet_data = pd.read_excel(excel_file, sheet_name=sheet_name)

            # Find header row index containing 'IP'
            header_row_index = None
            for i, row in sheet_data.iterrows():
                if 'IP' in row.values:
                    header_row_index = i
                    break
            
            if header_row_index is not None:
                # Set the header row and drop rows above it
                sheet_data.columns = sheet_data.iloc[header_row_index]
                sheet_data = sheet_data.drop(range(header_row_index + 1)).reset_index(drop=True)

                # Drop rows where 'Title' column is empty
                if 'Title' in sheet_data.columns:
                    sheet_data = sheet_data.dropna(subset=['Title'])

                processed_sheets[sheet_name] = sheet_data
            else:
                # If no header row found, you might skip or keep raw data as is
                # Here, we keep the sheet unmodified
                processed_sheets[sheet_name] = sheet_data

        # Save the processed sheets back to the same file, excluding "Overview" and "Status Description"
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for sheet_name, df in processed_sheets.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print(f"Processed and saved: {filepath}")

    except Exception as e:
        print(f"Error processing {filepath}: {e}")

# Iterate over the files in the directory
for filename in os.listdir(directory):
    if filename.endswith('.xlsx'):
        print(f"Processing file: {filename}")
        process_excel_file(os.path.join(directory, filename))
