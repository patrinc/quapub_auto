import shutil
import os

# Base path where monthly folders and template folder reside
base_path = r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"

# Name of the new monthly folder to create
month_folder_new = "25_10 Monthly Scans"

# Path for the new monthly folder
new_folder_path = os.path.join(base_path, month_folder_new)

# Path to the template folder with full structure and Python scripts
template_folder = os.path.join(base_path, "template_structure_with_scripts")

# Create the new folder if it does not exist
os.makedirs(new_folder_path, exist_ok=True)

# Recursively copy everything from template folder to new monthly folder
shutil.copytree(template_folder, new_folder_path, dirs_exist_ok=True)

print(f"New monthly folder '{month_folder_new}' created with full structure and scripts from template.")
