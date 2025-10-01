from datetime import date
import os
from pathlib import Path
from typing import Dict, List

# Year and month configuration
year = "2025"
current_month_num = "10"
previous_month_num = "09"
previous_month = "September"
current_month = "October"
fixed_start_date = date(2025, 10, 1)
assessment_date = "01-10-25"

# Centralized base paths - UPDATE THESE FOR YOUR ENVIRONMENT
BASE_PATHS = {
    'onedrive_base': Path(r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"),
    'raw_data_base': Path(r"C:\Users\example.user1\OneDrive - exampledomain\General - GSRC Global\CIC\Qualys\Qualys Historic Raw Data"),
    'sharepoint_base': Path(r"C:\Users\example.user1\OneDrive - exampledomain\General - GSRC Global"),
    'backup_base': Path("backups"),  # Relative to project root
    'logs_base': Path("logs")        # Relative to project root
}

# Business constants
COUNTRIES = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
REPORT_TYPES = ['Network', 'Server']
KEY_COLUMNS = ['IP', 'QID', 'Port']

# SLA Configuration
SLA_DAYS = {
    5: 7,   # Critical: 7 days
    4: 15,  # High: 15 days
    3: 21,  # Medium: 21 days
    2: 30,  # Low: 30 days
    1: 30   # Info: 30 days
}

NON_BREACH_STATUSES = {
    "Patched", "Mitigated", "False Positive", 
    "Not Applicable", "Risk Accepted", "Inactive", "Shutdown"
}

BREACH_STATUSES = {"Unpatched", "In Progress", "Deferred", ""}

# Filtering criteria for data processing
FILTER_CRITERIA = {
    'remove_empty_solution': True,
    'remove_types': ['Ig', 'Practice'],
    'remove_severity': [1],  # Remove Severity 1 (Info level)
    'default_status': 'Reviewed'
}

# Column configurations
COLUMNS_TO_DELETE_WITHOUT_OVERVIEW = [
    'Comments', 'Status_prev', 'Ticket No._prev', 'Comments_prev',
    'IT_prev', 'SLA Status_prev', 'composite_key_no_port'
]

# Dynamic path builders
def get_monthly_folder() -> str:
    """Get current monthly folder name"""
    return f"{year[-2:]}_{current_month_num} Monthly Scans"

def get_previous_monthly_folder() -> str:
    """Get previous monthly folder name"""
    return f"{year[-2:]}_{previous_month_num} Monthly Scans"

def get_raw_reports_folder() -> str:
    """Get raw reports folder name"""
    return f"{year}_{current_month_num}_Raw_Reports"

def get_monthly_path() -> Path:
    """Get path to current monthly folder"""
    return BASE_PATHS['onedrive_base'] / get_monthly_folder()

def get_previous_monthly_path() -> Path:
    """Get path to previous monthly folder"""
    return BASE_PATHS['onedrive_base'] / get_previous_monthly_folder()

def get_raw_data_path() -> Path:
    """Get path to raw data folder"""
    return BASE_PATHS['raw_data_base'] / get_raw_reports_folder()

def get_subfolder_path(subfolder: str) -> Path:
    """Get path to a subfolder within monthly folder"""
    return get_monthly_path() / subfolder

# Commonly used paths
def get_status_overview_path() -> Path:
    """Get path to 'with Status and Overview' folder"""
    return get_subfolder_path("01_with_Status_and_Overview")

def get_without_overview_path() -> Path:
    """Get path to 'without Overview' folder"""
    return get_subfolder_path("02_Final Reports_without_Overview")

def get_summary_requested_path() -> Path:
    """Get path to 'Summary Requested' folder"""
    return get_subfolder_path("03_Summary Requested")

def get_previous_reports_path() -> Path:
    """Get path to previous reports folder"""
    return get_subfolder_path("00_Previous_Final_Reports_And_Exceptions")

def get_sla_status_path() -> Path:
    """Get path to SLA Status folder"""
    return get_status_overview_path() / "SLA Status"

def get_metadata_path() -> Path:
    """Get path to metadata folder"""
    return get_status_overview_path() / "All Metadata"

def get_evidence_path() -> Path:
    """Get path to evidence folder"""
    return get_monthly_path() / "Monthly_Evidences"

# Country-specific SLA folders
COUNTRY_SLA_FOLDERS = {
    "Brazil": "Brazil Qualys Monthly Scans",
    "CSA": "CSA Qualys Monthly Scans", 
    "India": "India Qualys Monthly Scans",
    "Mexico": "Mexico Qualys Monthly Scans",
    "SFTL": "SFTL Qualys Monthly Scans"
}

def get_country_sla_path(country: str) -> Path:
    """Get SLA path for specific country"""
    folder_name = COUNTRY_SLA_FOLDERS.get(country, f"{country} Qualys Monthly Scans")
    return get_sla_status_path() / folder_name

# File naming helpers
def build_raw_report_filename(country: str, report_type: str) -> str:
    """Build raw report filename"""
    return f"{country}_{year}_{current_month_num}_Raw_Report_{report_type}_Report.xlsx"

def build_final_report_filename(country: str, report_type: str, with_status: bool = False) -> str:
    """Build final report filename"""
    base = f"{country}_{year}_{current_month_num}_Final_Report"
    if with_status:
        base += "_with_Status"
    base += f"_{report_type}_Report.xlsx"
    return base

def build_metadata_filename(bu: str, report_type: str) -> str:
    """Build metadata filename"""
    return f"{bu}_{year}_{current_month_num}_Metadata_{report_type}_Report.xlsx"

# Validation helpers
def validate_country(country: str) -> bool:
    """Validate country name"""
    return country in COUNTRIES

def validate_report_type(report_type: str) -> bool:
    """Validate report type"""
    return report_type in REPORT_TYPES

def get_required_paths_for_script(script_name: str) -> List[Path]:
    """Get required paths for a specific script"""
    # This maps script names to their required paths
    path_requirements = {
        '00_raw_report_network_segregation.py': [get_raw_data_path()],
        '00_raw_report_server_segregation.py': [get_raw_data_path()],
        '06_process_and_update_reports.py': [get_monthly_path()],
        '07_exception_removal.py': [get_status_overview_path(), get_previous_reports_path()],
        '08_copy_to_without_overview_folder.py': [get_status_overview_path()],
        '15_overdue_calc.py': [get_sla_status_path()],
        '18_upload_final_reports_with_status_internal_sharepoint.py': [get_without_overview_path()],
        '20_summary_ondemand.py': [get_without_overview_path()],
    }
    
    return path_requirements.get(script_name, [get_monthly_path()])

# Configuration validation
def validate_configuration() -> Dict[str, bool]:
    """Validate that all configured paths exist"""
    results = {}
    
    # Check base paths
    for name, path in BASE_PATHS.items():
        if name in ['backup_base', 'logs_base']:
            # These are created automatically
            results[f"base_path_{name}"] = True
        else:
            results[f"base_path_{name}"] = path.exists()
    
    # Check monthly paths (these might not exist yet)
    results['monthly_path'] = get_monthly_path().exists()
    results['raw_data_path'] = get_raw_data_path().exists()
    
    return results

def print_configuration_summary():
    """Print configuration summary"""
    print("\n" + "="*60)
    print("📋 QUALYS AUTOMATION CONFIGURATION")
    print("="*60)
    print(f"Year: {year}")
    print(f"Current Month: {current_month} ({current_month_num})")
    print(f"Previous Month: {previous_month} ({previous_month_num})")
    print(f"Assessment Date: {assessment_date}")
    print(f"Monthly Folder: {get_monthly_folder()}")
    print()
    print("📁 Key Paths:")
    print(f"  Monthly Folder: {get_monthly_path()}")
    print(f"  Raw Data: {get_raw_data_path()}")
    print(f"  Status & Overview: {get_status_overview_path()}")
    print(f"  Without Overview: {get_without_overview_path()}")
    print()
    print("🌍 Countries:", ", ".join(COUNTRIES))
    print("📊 Report Types:", ", ".join(REPORT_TYPES))
    print("="*60)

# Environment-specific overrides
def load_environment_overrides():
    """Load environment-specific configuration overrides"""
    env = os.getenv('QUALYS_ENV', 'production').lower()
    
    if env == 'development':
        # Development overrides
        pass
    elif env == 'testing':
        # Testing overrides - might use different paths
        pass
    
    # Load from environment variables if present
    if 'QUALYS_ONEDRIVE_BASE' in os.environ:
        BASE_PATHS['onedrive_base'] = Path(os.environ['QUALYS_ONEDRIVE_BASE'])
    
    if 'QUALYS_RAW_DATA_BASE' in os.environ:
        BASE_PATHS['raw_data_base'] = Path(os.environ['QUALYS_RAW_DATA_BASE'])

# Load overrides on import
load_environment_overrides()

# Backward compatibility - expose original variable names
base_path = str(BASE_PATHS['onedrive_base'])
countries = COUNTRIES
report_types = REPORT_TYPES
key_cols = KEY_COLUMNS 