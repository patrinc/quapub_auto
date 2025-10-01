import os

def rename_files(action, directory='.'):
    for file_name in os.listdir(directory):
        if file_name.endswith('.xlsx'):
            try:
                name, ext = os.path.splitext(file_name)
                parts = name.split('_')

                old_path = os.path.join(directory, file_name)

                if action == 'remove' and '_with_Status' in file_name:
                    new_name = file_name.replace('_with_Status', '')
                    new_path = os.path.join(directory, new_name)
                    os.rename(old_path, new_path)
                    print(f"Renamed: {file_name} -> {new_name}")

                elif action == 'add' and '_with_Status' not in file_name:
                    # Check if 'Final_Report' is at 4th position (index 3)
                    if len(parts) >= 5 and parts[3] == 'Final' and parts[4] == 'Report':
                        # Insert 'with_Status' after 'Final_Report'
                        # parts: [Country, Year, Month, Final, Report, ReportType...]
                        # We'll join 'Final' and 'Report' and add '_with_Status' after it
                        new_name = '_'.join(parts[:5]) + '_with_Status' + '_' + '_'.join(parts[5:]) + ext
                        new_path = os.path.join(directory, new_name)
                        os.rename(old_path, new_path)
                        print(f"Renamed: {file_name} -> {new_name}")
                    else:
                        print(f"Skipped (filename format unexpected): {file_name}")

            except Exception as e:
                print(f"Error renaming file {file_name}: {e}")

# Example usage:
rename_files('remove')  # To remove '_with_Status'
# rename_files('add')     # To add '_with_Status'
