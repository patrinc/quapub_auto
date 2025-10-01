from pathlib import Path
from datetime import datetime

# Base directory containing the BU folders
base_dir = Path(r"C:\Users\example.user1\OneDrive - exampledomain\Qualys Monthly Scans")

# Get current year and month for naming the monthly folder
now = datetime.now()
monthly_folder_name = f"{now.year}_{now.month:02d}"  # e.g., "2025_06"

# List of BU folders (you can also dynamically get these if needed)
bu_folders = [
    "Brazil Qualys Monthly Scans",
    "CSA Qualys Monthly Scans",
    "India Qualys Monthly Scans",
    "Mexico Qualys Monthly Scans",
    "SFTL Qualys Monthly Scans"
]

for bu in bu_folders:
    bu_path = base_dir / bu
    evidence_path = bu_path / "Monthly_Evidences"
    monthly_path = evidence_path / monthly_folder_name
    
    # Create evidence/monthly folder structure
    monthly_path.mkdir(parents=True, exist_ok=True)
    
    print(f"Created or verified folder: {monthly_path}")
