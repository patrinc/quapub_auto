import re
import os
from prettytable import PrettyTable
from colorama import Fore, Style, init as colorama_init
from rich.console import Console
from rich.table import Table
from rich import box

# Initialize Colorama
colorama_init(autoreset=True)

class UnifiedFormatter:
    def __init__(self):
        self.console = Console()

    def format_line(self, line, line_type="normal", indent=0, width=70):
        indent_str = " " * indent
        if line_type == "title":
            self.console.print(f"\n=== {line} ===", style="cyan")
        elif line_type == "header":
            print(Fore.YELLOW + indent_str + line + Style.RESET_ALL)
        elif line_type == "processing":
            print(Fore.BLUE + indent_str + line + Style.RESET_ALL)
        elif line_type == "success":
            print(Fore.GREEN + indent_str + line + Style.RESET_ALL)
        elif line_type == "summary":
            self.console.print(line.center(width), style="magenta")
        elif line_type == "error_title":
            print(Fore.RED + indent_str + line + Style.RESET_ALL)
        elif line_type == "error_line":
            print(Fore.RED + indent_str + line + Style.RESET_ALL)
        else:
            print(indent_str + line)

    def print_table(self, title, columns, rows, styles=None, justify=None, box_style=box.MINIMAL_DOUBLE_HEAD):
        table = Table(title=title, box=box_style)
        for i, col in enumerate(columns):
            style = styles[i] if styles and i < len(styles) else None
            just = justify[i] if justify and i < len(justify) else None
            table.add_column(col, style=style, justify=just)
        for row in rows:
            table.add_row(*[str(item) for item in row])
        self.console.print(table)

# Create a singleton formatter instance
formatter = UnifiedFormatter()


def parse_network_segregation(stdout, stderr):
    formatter.format_line("Network Segregation Script Output", "title")
    for line in stdout.splitlines():
        if "Processing file:" in line:
            formatter.format_line(line, "processing", indent=2)
        elif "Saved metadata Excel:" in line or "Saved data Excel:" in line:
            formatter.format_line(line, "success", indent=4)
        elif "All network reports and metadata" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_server_segregation(stdout, stderr):
    formatter.format_line("Server Segregation Script Output", "title")
    for line in stdout.splitlines():
        if "Processing file:" in line:
            formatter.format_line(line, "processing", indent=2)
        elif "Saved metadata Excel:" in line or "Saved data Excel:" in line:
            formatter.format_line(line, "success", indent=4)
        elif "All server reports and metadata" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_sharepoint_report(stdout, stderr):
    formatter.format_line("SharePoint Report Retrieval Summary", "title")

    copy_entries = {}
    script_running_line = ""
    script_finished_line = ""

    for line in stdout.splitlines():
        line = line.strip()
        if line.startswith("Ready to run"):
            script_running_line = line
        elif line.startswith("Copied:"):
            parts = line[len("Copied: "):].split(" -> ")
            if len(parts) == 2:
                src = parts[0]
                filename = src.split("\\")[-1]  # Fixed escaping for backslash

                tokens = filename.split("_")
                if len(tokens) >= 7:
                    region = tokens[0]
                    report_type = tokens[-2] + " " + tokens[-1].replace(".xlsx", "")
                else:
                    region = "Unknown"
                    report_type = filename

                if region not in copy_entries:
                    copy_entries[region] = {}
                copy_entries[region][report_type] = filename
        elif "Finished running" in line:
            script_finished_line = line

    if script_running_line:
        formatter.format_line(script_running_line + "\n")

    if copy_entries:
        table = PrettyTable()
        table.field_names = ["Region", "Network Report Filename", "Server Report Filename"]

        # Align columns for neatness
        table.align["Region"] = "l"
        table.align["Network Report Filename"] = "l"
        table.align["Server Report Filename"] = "l"

        for region in sorted(copy_entries.keys()):
            network_report = copy_entries[region].get("Network Report", "")
            server_report = copy_entries[region].get("Server Report", "")

            # Fallback to keys containing "Network" or "Server" if exact keys absent
            if not network_report:
                for k in copy_entries[region]:
                    if "Network" in k:
                        network_report = copy_entries[region][k]
            if not server_report:
                for k in copy_entries[region]:
                    if "Server" in k:
                        server_report = copy_entries[region][k]

            table.add_row([region, network_report, server_report])

        # Print table lines with indentation and 'success' style using formatter
        for line in table.get_string().splitlines():
            formatter.format_line(line, "success", indent=2)

        formatter.format_line("\nAll files were copied to the destination folder.\n")

    if script_finished_line:
        total_width = max(len(line) for line in table.get_string().splitlines()) if copy_entries else 0
        formatter.format_line(script_finished_line, "summary", total_width)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_delete_overview(stdout, stderr):
    formatter.format_line("Delete Overview Script Output", "title")
    for line in stdout.splitlines():
        if "Processing file:" in line:
            formatter.format_line(line, "processing", indent=2)
        elif "Processed and saved:" in line:
            formatter.format_line(line, "success", indent=4)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_status_color(stdout, stderr):
    formatter.format_line("Status Color Script Output", "title")
    for line in stdout.splitlines():
        if "Processing file:" in line:
            formatter.format_line(line, "processing", indent=2)
        elif "Applied colors and saved:" in line:
            formatter.format_line(line, "success", indent=4)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_exception_list(stdout, stderr):
    formatter.format_line("Exception List Script Output Summary", "title")
    script_running_line = ""
    counts = {}
    save_path = ""
    script_finished_line = ""

    for line in stdout.splitlines():
        line = line.strip()
        if line.startswith("Ready to run"):
            script_running_line = line
        elif "count:" in line:
            parts = line.split("count:")
            if len(parts) == 2:
                desc = parts[0].strip()
                count = parts[1].strip()
                counts[desc] = count
        elif line.startswith("Saved both lists with colors to"):
            save_path = line[len("Saved both lists with colors to"):].strip()
        elif "Finished running" in line:
            script_finished_line = line

    if script_running_line:
        formatter.format_line(script_running_line + "\n")

    if counts:
        table = PrettyTable()
        table.field_names = ["Description", "Count"]

        # Align columns nicely
        table.align["Description"] = "l"
        table.align["Count"] = "r"

        # Add all count pairs as rows
        for desc, count in counts.items():
            table.add_row([desc, count])

        # Print table lines with indentation and success style
        for line in table.get_string().splitlines():
            formatter.format_line(line, "success", indent=2)

    if save_path:
        formatter.format_line("\nSaved colored lists to:")
        formatter.format_line(save_path, "success", indent=2)

    if script_finished_line:
        total_width = max(70, max((len(desc) for desc in counts.keys()), default=0) + 12)
        formatter.format_line(script_finished_line, "summary", total_width)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_change_name_reports(stdout, stderr):
    formatter.format_line("Rename Reports Script Output", "title")
    for line in stdout.splitlines():
        if "Files have been renamed successfully." in line:
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_process_and_update_reports(stdout, stderr):
    formatter.format_line("Process and Update Reports Script Output", "title")

    table = Table(title="", box=box.MINIMAL_DOUBLE_HEAD)
    table.add_column("Country - Report", style="blue")
    table.add_column("Rows Loaded", justify="right", style="yellow")
    table.add_column("Removed empty 'Solution'", justify="right", style="red")
    table.add_column("Removed Type 'Ig/Practice'", justify="right", style="red")
    table.add_column("Removed Severity 1", justify="right", style="red")
    table.add_column("Final Report Path", style="green")

    current_report = None
    rows_loaded = 0
    removed_empty = 0
    removed_type = 0
    removed_severity = 0
    final_report_path = ""

    def add_report_row():
        nonlocal current_report, rows_loaded, removed_empty, removed_type, removed_severity, final_report_path
        if current_report is not None:
            table.add_row(
                current_report,
                str(rows_loaded),
                str(removed_empty),
                str(removed_type),
                str(removed_severity),
                final_report_path
            )
        current_report = None
        rows_loaded = 0
        removed_empty = 0
        removed_type = 0
        removed_severity = 0
        final_report_path = ""

    for line in stdout.splitlines():
        if line.startswith("Processing"):
            add_report_row()
            current_report = line.replace("Processing ", "").replace(" reports...", "")
            formatter.format_line(line, "processing")
        elif line.startswith("Loaded file:"):
            parts = line.split("| Rows:")
            if len(parts) == 2:
                try:
                    rows_loaded = int(parts[1].strip())
                except ValueError:
                    rows_loaded = 0
            formatter.format_line("  " + line, "success")
        elif "Removed" in line:
            match = re.search(r'\b(\d+)\b', line)
            if match:
                number = int(match.group(1))
                if "empty 'Solution'" in line:
                    removed_empty = number
                elif "Type 'Ig' or 'Practice'" in line:
                    removed_type = number
                elif "Severity 1" in line:
                    removed_severity = number
        elif line.startswith("[") and "Final report saved:" in line:
            final_report_path = line.split(": ", 1)[1]
            formatter.format_line("  " + line, "success")
        elif "All reports processed successfully." in line or line.startswith("Finished running"):
            add_report_row()
            formatter.format_line(line.center(70), "summary")
        else:
            formatter.format_line(line)

    formatter.console.print(table)

    if stderr.strip():
        formatter.format_line("\nErrors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_exception_removal(stdout, stderr):
    total_loaded = None
    removals = {}
    in_summary_section = False

    for line in stdout.splitlines():
        if "Loaded" in line:
            total_loaded = int(line.split()[1])
        elif line.startswith("Summary of removals per file:"):
            in_summary_section = True
        elif in_summary_section and "rows removed" in line:
            parts = line.split(":")
            filename = parts[0].strip()
            removed_rows = int(parts[1].strip().split()[0])
            removals[filename] = removed_rows

    # Title and summary header lines
    formatter.format_line("Exception Removal Summary", "title")
    formatter.format_line(f"Total exceptions loaded from previous month: {total_loaded}", "header", indent=2)
    formatter.format_line(f"Total files processed: {len(removals)}", "header", indent=2)
    formatter.format_line(f"Total rows removed: {sum(removals.values())}\n", "header", indent=2)

    # Create PrettyTable with aligned columns and borders
    table = PrettyTable()
    table.field_names = ["Filename", "Rows Removed"]

    # Align Filename left, Rows Removed right
    table.align["Filename"] = "l"
    table.align["Rows Removed"] = "r"

    # Add all data rows
    for filename, removed in removals.items():
        table.add_row([filename, removed])

    # Format each table line with indentation and success style
    for line in table.get_string().splitlines():
        formatter.format_line(line, "success", indent=2)

    # Print stderr errors if any
    if stderr.strip():
        formatter.format_line("\nErrors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_copy_without_overview(stdout, stderr):
    formatter.format_line("Copy to Without Overview Folder Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Found matching file:"):
            formatter.format_line(line, "processing", indent=2)  # Blue color for info
        elif line.startswith("Copied and renamed file to:"):
            formatter.format_line(line, "success", indent=2)  # Green color for success
        elif line.startswith("Deleted specified columns"):
            formatter.format_line(line, "header", indent=2)  # Yellow color for header/notice
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)  # Centered magenta for finish
        else:
            formatter.format_line(line)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_insert_metadata(stdout, stderr):
    formatter.format_line("Insert Metadata Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Processing final report:"):
            formatter.format_line(line, "processing")
        elif line.startswith("Using metadata file:") or line.startswith("Metadata size:"):
            formatter.format_line(line, "header", indent=2)
        elif line.startswith("Metadata inserted successfully"):
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_create_individual_overviews(stdout, stderr):
    formatter.format_line("Create Individual Overviews Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("File") and "is ready for overview creation." in line:
            formatter.format_line(line, "processing", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_generate_summary(stdout, stderr):
    formatter.format_line("Generate Summary Script Output", "title")
    for line in stdout.splitlines():
        if "Summary report saved to" in line:
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_rename(stdout, stderr):
    formatter.format_line("Rename Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Renamed:"):
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_script_output(stdout, stderr):
    formatter.format_line("Script Output", "title")
    for line in stdout.splitlines():
        if "Processing" in line:
            formatter.format_line(line, "processing")
        elif any(keyword in line for keyword in ["Saved", "Copied", "Renamed", "Applied"]):
            formatter.format_line(line, "success")
        elif any(keyword in line for keyword in ["Removed", "Deleted", "Exception"]):
            formatter.format_line(line, "error_line")
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line")

def parse_copy_to_sla_status(stdout, stderr):
    formatter.console.print("\n=== Copy to SLA Status Script Output ===", style="cyan")

    source_folder = None
    destination_base = None
    copied_files = []

    for line in stdout.splitlines():
        if line.startswith("Starting file copy for Year:"):
            formatter.console.print(line, style="blue")
        elif line.startswith("Source folder:"):
            source_folder = line.split("Source folder:")[1].strip()
            formatter.console.print(line, style="yellow")
        elif line.startswith("Destination base path:"):
            destination_base = line.split(":")[1].strip()
            formatter.console.print(line, style="yellow")
        elif line.startswith("Copied '"):
            parts = line.split("'")
            if len(parts) >= 4:
                file_name = parts[1]
                dest_path = parts[3]
                copied_files.append((file_name, dest_path))
        elif line == "File copy completed." or line.startswith("Finished running"):
            formatter.console.print(line.center(70), style="magenta")
        else:
            formatter.console.print(line)

    if copied_files:
        table = Table(title="", box=box.MINIMAL_DOUBLE_HEAD)
        table.add_column("File Name", style="cyan")
        table.add_column("Destination Folder", style="green")

        for file_name, dest_path in copied_files:
            table.add_row(file_name, dest_path)

        formatter.console.print(table)

    if stderr.strip():
        formatter.console.print("\nErrors:", style="red")
        for line in stderr.splitlines():
            formatter.console.print("  " + line, style="red")

def parse_overdue_calc(stdout, stderr):
    formatter.format_line("Overdue Calculation Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Processing:"):
            formatter.format_line(line, "processing")
        elif line.startswith("Elapsed calendar days"):
            formatter.format_line(line, "header", indent=2)
        elif line.startswith("Severity levels currently overdue"):
            formatter.format_line(line, "error_line", indent=2)
        elif line.startswith("SLA breaches identified"):
            count = int(line.split(":")[-1].strip())
            color_type = "success" if count == 0 else "error_line"
            formatter.format_line(line, color_type, indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)

    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_evidence_folder(stdout, stderr):
    formatter.format_line("Evidence Folder Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Created or verified folder:"):
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_old_new_status(stdout, stderr):
    formatter.console.print("\n=== Old New Status Update Script Output ===", style="cyan")

    updated_files = []
    for line in stdout.splitlines():
        if line.startswith("Updated 'Status' column in file:"):
            full_path = line.split("file:")[1].strip()
            filename = os.path.basename(full_path)
            updated_files.append(filename)
        elif "Finished running" in line:
            formatter.console.print(line.center(70), style="magenta")
        else:
            formatter.console.print(line)

    if updated_files:
        table = PrettyTable()
        table.field_names = ["File Name", "Status"]

        # Align columns: Filename left, Status center
        table.align["File Name"] = "l"
        table.align["Status"] = "c"

        for fname in updated_files:
            table.add_row([fname, "Updated"])

        # Print table lines with indentation and success style
        for line in table.get_string().splitlines():
            formatter.format_line(line, "success", indent=2)

        formatter.console.print(f"\nTotal files updated: {len(updated_files)}\n", style="bold yellow")

    if stderr.strip():
        formatter.console.print("\nErrors:", style="red")
        for line in stderr.splitlines():
            formatter.console.print("  " + line, style="red")

def parse_upload_final_reports_sharepoint(stdout, stderr):
    formatter.console.print("\n=== Upload Final Reports to SharePoint Script Output ===", style="cyan")
    source_folder = None
    destination_base = None
    copied_files = []

    for line in stdout.splitlines():
        if line.startswith("Scanning for files in:"):
            source_folder = line.split("Scanning for files in:")[1].strip()
            formatter.console.print(f"Scanning for files in: {source_folder}", style="blue")
        elif line.startswith("Copying files to:"):
            destination_base = line.split("Copying files to:")[1].strip()
            formatter.console.print(f"Copying files to: {destination_base}", style="blue")
        elif line.startswith("Copied file:"):
            parts = line.split("Copied file:")[1].strip().split(" to ")
            if len(parts) == 2:
                filename = os.path.basename(parts[0])
                dest_path = parts[1]
                relative_dest = dest_path[len(destination_base):].lstrip("\\/") if destination_base and dest_path.startswith(destination_base) else dest_path
                copied_files.append((filename, relative_dest))
        elif "Copy complete." in line or "Finished running" in line:
            formatter.console.print(line.center(70), style="magenta")
        else:
            formatter.console.print(line)

    if copied_files:
        table = Table(title="", box=box.MINIMAL_DOUBLE_HEAD)
        table.add_column("File Name", style="cyan")
        table.add_column("Destination (Relative Path)", style="green")
        for filename, rel_dest in copied_files:
            table.add_row(filename, rel_dest)
        formatter.console.print(table)
        formatter.console.print(f"Total files copied: {len(copied_files)}\n", style="bold yellow")

    if stderr.strip():
        formatter.console.print("\nErrors:", style="red")
        for line in stderr.splitlines():
            formatter.console.print("  " + line, style="red")

def parse_summary_ondemand(stdout, stderr):
    formatter.format_line("Summary On-demand Script Output", "title")
    for line in stdout.splitlines():
        if "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line")

def parse_export_charts(stdout, stderr):
    formatter.console.print("\n=== Export Charts Script Output ===", style="cyan")
    exported_charts = []
    for line in stdout.splitlines():
        if line.startswith("Exported chart from sheet"):
            parts = line.split("'")
            if len(parts) >= 3:
                sheet_name = parts[1]
                file_path_part = line.split(" as ")[-1].strip()
                filename = os.path.basename(file_path_part)
                exported_charts.append((sheet_name, filename))
        elif "Finished running" in line:
            formatter.console.print(line.center(70), style="magenta")
        else:
            formatter.console.print(line)

    if exported_charts:
        table = PrettyTable()
        table.field_names = ["Sheet Name", "File Name"]

        # Alignment for readability
        table.align["Sheet Name"] = "l"
        table.align["File Name"] = "l"

        for sheet_name, filename in exported_charts:
            table.add_row([sheet_name, filename])

        # Print table line-by-line with indentation and formatter styles
        for line in table.get_string().splitlines():
            formatter.format_line(line, "success", indent=2)

        formatter.console.print(f"\nTotal charts exported: {len(exported_charts)}", style="bold yellow")

    if stderr.strip():
        formatter.console.print("\nErrors:", style="red")
        for line in stderr.splitlines():
            formatter.console.print("  " + line, style="red")

def parse_draft_internal_email(stdout, stderr):
    formatter.format_line("Draft Internal Email Script Output", "title")
    for line in stdout.splitlines():
        if "Summary table extracted and styled successfully." in line:
            formatter.format_line(line, "success", indent=2)
        elif "Attached" in line:
            formatter.format_line(line, "header", indent=2)
        elif "Email created and displayed successfully with" in line:
            formatter.format_line(line, "success", indent=2)
        elif "Please review and send manually." in line:
            formatter.format_line(line, "summary", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)

def parse_reply_vapt_emails(stdout, stderr):
    formatter.format_line("Reply VAPT Emails Script Output", "title")
    for line in stdout.splitlines():
        if line.startswith("Replying to:"):
            formatter.format_line(line, "processing", indent=2)
        elif line.startswith("Flagged emails matched:") or line.startswith("Replies processed:"):
            formatter.format_line(line, "success", indent=2)
        elif "Finished running" in line:
            formatter.format_line(line, "summary", width=70)
        else:
            formatter.format_line(line)
    if stderr.strip():
        formatter.format_line("Errors:", "error_title")
        for line in stderr.splitlines():
            formatter.format_line(line, "error_line", indent=2)
