import os

# Define the directory containing the files
directory = os.getcwd()

# Loop through all files in the directory
for filename in os.listdir(directory):
    if filename.endswith(".xlsx"):  # Check if the file is an Excel file
        # Replace "Reports" with "Report" or vice versa
        new_filename = filename.replace("Reports", "Report")  # Change to "Report"
        # new_filename = filename.replace("Report", "Reports")  # Uncomment this line to change to "Reports"
        
        # Rename the file
        os.rename(os.path.join(directory, filename), os.path.join(directory, new_filename))

print("Files have been renamed successfully.")
