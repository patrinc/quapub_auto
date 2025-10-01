import os
import re
import subprocess
from colorama import init, Fore
import custom_parsers

init(autoreset=True)

def get_latest_monthly_folder(parent_dir="."):
    """
    Find the latest folder matching pattern: 'YY_MM Monthly Scans' inside parent_dir
    """
    pattern = re.compile(r"(\d{2})_(\d{2}) Monthly Scans")
    folders = [f for f in os.listdir(parent_dir) if os.path.isdir(os.path.join(parent_dir, f))]
    matching_folders = []
    for folder in folders:
        match = pattern.match(folder)
        if match:
            year, month = int(match.group(1)), int(match.group(2))
            matching_folders.append((year, month, folder))
    if not matching_folders:
        print(Fore.RED + "No matching monthly scan folders found.")
        return None
    # Sort by year and month descending to get latest
    matching_folders.sort(key=lambda x: (x[0], x[1]), reverse=True)
    latest_folder = matching_folders[0][2]
    return os.path.join(parent_dir, latest_folder)

def extract_number(filename):
    """
    Extract leading number from filename for sorting.
    """
    match = re.match(r"(\d+)", filename)
    return int(match.group(1)) if match else float('inf')

def gather_scripts(folder):
    """
    Recursively gather all Python (*.py) scripts inside folder.
    Returns list of tuples: (relative_path, full_path)
    """
    scripts = []
    for root, dirs, files in os.walk(folder):
        for file in files:
            if file.endswith(".py"):
                full_path = os.path.join(root, file)
                rel_path = os.path.relpath(full_path, folder)
                scripts.append((rel_path, full_path))
    return scripts

def parse_delete_status(stdout, stderr):
    """
    Fallback parser for 13_delete_status.py if missing in custom_parsers.
    """
    print(Fore.CYAN + "\n=== Delete Status Script Output (Fallback) ===")
    if stdout:
        print(Fore.RESET + stdout)
    if stderr:
        print(Fore.RED + "Errors:\n" + stderr)

SCRIPT_PARSER_MAP = {
    "00_raw_report_network_segregation.py": custom_parsers.parse_network_segregation,
    "00_raw_report_server_segregation.py": custom_parsers.parse_server_segregation,
    "01_get_all_reports_from_sharepoint_latest.py": custom_parsers.parse_sharepoint_report,
    "02_delete_overview.py": custom_parsers.parse_delete_overview,
    "03_status_color.py": custom_parsers.parse_status_color,
    "04_exception_list.py": custom_parsers.parse_exception_list,
    "05_change_name_reports_to_report.py": custom_parsers.parse_change_name_reports,
    "06_process_and_update_reports.py": custom_parsers.parse_process_and_update_reports,
    "07_exception_removal.py": custom_parsers.parse_exception_removal,
    "08_copy_to_without_overview_folder.py": custom_parsers.parse_copy_without_overview,
    "09_insert_metadata.py": custom_parsers.parse_insert_metadata,
    "10_create_individual_overviews_with_date.py": custom_parsers.parse_create_individual_overviews,
    "11_generate_summary.py": custom_parsers.parse_generate_summary,
    "12_rename.py": custom_parsers.parse_rename,
    "13_delete_status.py": getattr(custom_parsers, "parse_delete_status", parse_delete_status),
    "14_copy_to_sla_status.py": custom_parsers.parse_copy_to_sla_status,
    "15_overdue_calc.py": custom_parsers.parse_overdue_calc,
    "16_evidence_folder.py": custom_parsers.parse_evidence_folder,
    "17_old_new_status.py": custom_parsers.parse_old_new_status,
    "18_upload_final_reports_with_status_internal_sharepoint.py": custom_parsers.parse_upload_final_reports_sharepoint,
    "20_summary_ondemand.py": custom_parsers.parse_summary_ondemand,
    "21_export_charts.py": custom_parsers.parse_export_charts,
    "22_draft_internal_email.py": custom_parsers.parse_draft_internal_email,
    "23_reply_vapt_emails.py": custom_parsers.parse_reply_vapt_emails,
}

def run_script(rel_path, full_path):
    script_dir = os.path.dirname(full_path)
    script_name = os.path.basename(full_path)

    input(Fore.CYAN + f"Ready to run {rel_path}? Press Enter to continue or Ctrl+C to abort...")
    print(Fore.CYAN + f"Running {rel_path} ...")

    result = subprocess.run(
        ["python", script_name],
        cwd=script_dir,
        capture_output=True,
        text=True,
        shell=False
    )

    parser = SCRIPT_PARSER_MAP.get(script_name)
    if parser:
        parser(result.stdout, result.stderr)
    else:
        if result.stdout:
            print(Fore.RESET + result.stdout)
        if result.stderr:
            print(Fore.RED + "Errors:\n" + result.stderr)

    if result.returncode == 0:
        print(Fore.GREEN + f"Finished running {rel_path} successfully.\n{'-'*70}")
        return True
    else:
        print(Fore.RED + f"Script {rel_path} failed with exit code {result.returncode}. Stopping.")
        return False

def main():
    try:
        folder = get_latest_monthly_folder(".")
        if folder is None:
            return
        
        print(Fore.CYAN + f"Running scripts from latest folder: {folder}")

        scripts = gather_scripts(folder)
        scripts.sort(key=lambda x: extract_number(os.path.basename(x[0])))

        print(Fore.CYAN + "Available scripts:")
        for idx, (rel_path, _) in enumerate(scripts, start=1):
            print(f"{idx}: {rel_path}")
        print('Enter "1" to start from the first script, or any script number to start from there.')

        choice = input(Fore.CYAN + "Enter script number to start running from that script (Ctrl+C to abort anytime): ").strip()

        script_start_idx = 0
        try:
            script_start_idx = int(choice) - 1
            if not (0 <= script_start_idx < len(scripts)):
                print(Fore.RED + "Invalid script number. Starting from the beginning.")
                script_start_idx = 0
        except ValueError:
            print(Fore.RED + "Invalid input. Starting from the beginning.")
            script_start_idx = 0

        for rel_path, full_path in scripts[script_start_idx:]:
            if not run_script(rel_path, full_path):
                break

    except KeyboardInterrupt:
        print(Fore.YELLOW + "\nExecution aborted by user with Ctrl+C.")

if __name__ == "__main__":
    main()
