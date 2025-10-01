# Qualys Monthly Automation

Automated Python scripts to process, analyze, and report monthly Qualys vulnerability scan data across multiple countries and report types. Enables data cleansing, status updates, SLA calculation, summary generation, email drafting, and Outlook integration.

***

## Table of Contents

- [Project Overview](#project-overview)
- [Workflow and Script Descriptions](#workflow-and-script-descriptions)
- [Installation](#installation)
- [Usage](#usage)
- [Configuration](#configuration)
- [File and Folder Structure](#file-and-folder-structure)
- [Dependencies](#dependencies)
- [Contributing](#contributing)
- [Author and Contact](#author-and-contact)

***

## Project Overview

This project automates the end-to-end processing of monthly Qualys vulnerability scan reports for multiple Business Units (BUs), including:

- Standardizing and cleaning raw reports
- Removing exceptions from prior months
- Adding metadata and cleaning columns
- Copying files into organized monthly folders
- Calculating SLA breach status based on business rules
- Generating detailed summaries with Microsoft Excel charts
- Automating email composition and Outlook replies for approvals
- Exporting charts as PNGs for presentations
- Maintaining proper status tracking and response workflows

These scripts collectively reduce manual overhead, improve consistency, and accelerate vulnerability management communications.

***

## Workflow and Script Descriptions

Scripts are designed modularly, each numbered to represent sequence or logical step:

- `05_change_name_reports_to_report.py`: Rename files for naming consistency.
- `06_process_and_update_reports.py`: Clean and filter raw vulnerability reports; mark records as reviewed.
- `07_exception_removal.py`: Remove known exceptions from updated reports using previous month’s exception list.
- `08_copy_to_without_overview_folder.py`: Copy cleaned reports removing unnecessary columns.
- `09_insert_metadata.py`: Insert metadata rows and styles atop final reports.
- `10_create_individual_overviews_with_date.py`: Add overview sheets summarizing severity counts per report.
- `11_generate_summary.py`: Aggregate vulnerability data into Excel summary files with charts.
- `12_rename.py`: Batch rename reports to toggle '_with_Status' substring.
- `13_delete_status.py`: Add status-related columns, validations, legends, and prepare for approval.
- `14_copy_to_sla_status.py`: Copy status-tagged reports into monthly SLA tracker folders.
- `15_overdue_calc.py`: Calculate SLA breach status marking overdue vulnerabilities.
- `16_evidence_folder.py`: Create structured monthly evidence folders per BU.
- `17_old_new_status.py`: Compare current and previous month reports to label vulnerabilities as `New` or `Old`.
- `18_upload_final_reports_with_status_internal_sharepoint.py`: Copy final reports into an internal SharePoint synced folder.
- `20_summary_ondemand.py`: Dynamically generate consolidated on-demand vulnerability summaries with styled charts.
- `21_export_charts.py`: Programmatically export Excel charts as PNG images using Excel COM automation.
- `22_draft_internal_email.py`: Draft Outlook emails for approval requests embedding HTML summary tables and file attachments.
- `23_reply_vapt_emails.py`: Automatic Outlook reply to flagged emails with personalized summaries and SharePoint links.

***

## Installation

1. Clone the repo or download scripts to a local machine with Python 3.x installed.
2. Ensure the following Python packages are installed: `pandas`, `openpyxl`, `xlsxwriter`, `beautifulsoup4`, `pywin32`. Install via pip if needed:
```bash
pip install pandas openpyxl xlsxwriter beautifulsoup4 pywin32
```

3. Microsoft Excel and Outlook must be installed on the machine for COM automation scripts.

***

## Usage

1. Configure `config.py` for year, month, folder paths, and business-specific parameters.
2. Run scripts sequentially or as needed based on the workflow to process raw reports to final deliverables.
3. Review generated Excel summaries and email drafts before sending.
4. Use exported charts for internal meetings or presentations.
5. Monitor Outlook flagged emails to automate replies.

***

## Configuration

A central `config.py` holds shared parameters such as:

```python
from datetime import date

# Year and month configuration
year = "2025"
current_month_num = "10"
previous_month_num = "09"
previous_month = "September"
current_month = "October"

# Fixed start date as a datetime.date object
fixed_start_date = date(2025, 9, 29)

# Other config variables
assessment_date = "29-09-25"
...
```

Update paths, SLA settings, and email recipients as needed.

***

## File and Folder Structure

```
qualys-automation/
├── config.py                # Central configuration 
├── scripts/                 # Automation scripts by step order
│   ├── 05_change_name_reports_to_report.py
│   ├── 06_process_and_update_reports.py
│   ├── ...
│   └── 23_reply_vapt_emails.py
├── data/                    # Monthly report data organized by year_month
│   ├── 25_10 Monthly Scans/
│   │   ├── 00_Previous_Final_Reports_And_Exceptions
│   │   ├── 01_with_Status_and_Overview
│   │   ├── 02_Final Reports_without_Overview
│   │   ├── 03_Summary Requested
│   │   ├── Monthly_Evidences/
│   │   └── ...
├── outputs/                 # Generated summaries, charts, exported images
├── logs/                    # (Optional) logs of script runs
└── README.md                # Project documentation (this file)
```


***

## Dependencies

- Python 3.x
- pandas
- openpyxl
- xlsxwriter
- beautifulsoup4
- pywin32 (for Windows COM automation with Outlook and Excel)
- Microsoft Office Excel and Outlook installed

***

## Contributing

This project is maintained by GSRC Team.
Feel free to raise issues or submit pull requests.
Ensure tests and documentation accompany changes.

***
