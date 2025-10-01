import os
import sys
import pandas as pd
import glob
import xlsxwriter

# Add parent directory (where config.py is located) to sys.path to allow import
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Access configuration variables here

def process_reports():
    # Use config variables instead of hardcoded values
    previous_month_num = config.previous_month_num
    current_month_num = config.current_month_num
    year = config.year
    
    # Construct the dynamic folder path
    base_path = r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"
    folder_name = f"{year[-2:]}_{current_month_num} Monthly Scans"
    subfolder = "02_Final Reports_without_Overview"
    target_dir = os.path.join(base_path, folder_name, subfolder)
    countries = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
    report_types = ['Network', 'Server']

    files = glob.glob(os.path.join(target_dir, "*.xlsx"))  # Read files from the dynamic folder
    summary_data = {}

    for file in files:
        filename = os.path.basename(file)
        parts = filename.split("_")
        if len(parts) < 6:
            continue  # Skip files that don't match the expected naming structure
        bu_name = parts[0]
        report_type = parts[-2]
        if bu_name not in countries or report_type not in report_types:
            continue
        report_key = f"{report_type}_Report"
        if bu_name not in summary_data:
            summary_data[bu_name] = {
                "Network_Report": {severity: {"New": 0, "Old": 0} for severity in range(1, 6)},
                "Server_Report": {severity: {"New": 0, "Old": 0} for severity in range(1, 6)}
            }
        xls = pd.ExcelFile(file)
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, usecols=["Severity", "Status"], dtype={"Severity": str})
            df = df.dropna(subset=["Severity", "Status"])
            df["Status"] = df["Status"].astype(str)
            df["Severity"] = pd.to_numeric(df["Severity"].str.extract(r'(\d)')[0], errors='coerce')
            for severity in range(1, 6):
                summary_data[bu_name][report_key][severity]["New"] += (
                    df[(df["Severity"] == severity) & (df["Status"].str.startswith("New"))].shape[0]
                )
                summary_data[bu_name][report_key][severity]["Old"] += (
                    df[(df["Severity"] == severity) & (df["Status"] == "Old")].shape[0]
                )
    return summary_data

def save_summary_to_excel(summary_data, output_file):
    workbook = xlsxwriter.Workbook(output_file)

    # Header format matching your Outlook email table styles
    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#D9E1F2',  # Outlook header background color
        'font_color': '#1F497D',  # Outlook header text color
        'border': 1,
        'font_name': 'Calibri',
        'font_size': 11
    })

    # Fonts for chart titles and axes matching email style
    title_font = {'name': 'Calibri', 'color': '#1F497D', 'size': 14, 'bold': True}
    axis_font = {'name': 'Calibri', 'color': '#1F497D', 'size': 10}
    legend_font = {'name': 'Calibri', 'color': '#1F497D', 'size': 10}

    # Define colors for chart series: brand blue plus accents
    series_colors = {
        'New': '#4F81BD',      # Brand blue
        'Old': '#A9CCE3',      # Light blue accent
    }

    for bu, reports in summary_data.items():
        sheet_name = f"{bu}"[:31]  # Excel sheet name max length 31
        worksheet = workbook.add_worksheet(sheet_name)

        positions = {
            "Network_Report": (1, "B", "H2"),
            "Server_Report": (19, "B", "H20")
        }

        for report_type, (start_row, start_col_letter, chart_cell) in positions.items():
            start_col = ord(start_col_letter) - ord('A')
            headers = ["Severity", "New", "Old"]
            for col, header in enumerate(headers):
                worksheet.write(start_row, start_col + col, header, header_format)

            severity_labels = ["Info", "Low", "Medium", "High", "Critical"]
            for i, severity in enumerate(range(1, 6)):
                values = reports.get(report_type, {}).get(severity, {"New": 0, "Old": 0})
                worksheet.write(start_row + i + 1, start_col, severity_labels[i])
                worksheet.write_number(start_row + i + 1, start_col + 1, values["New"])
                worksheet.write_number(start_row + i + 1, start_col + 2, values["Old"])

            chart = workbook.add_chart({'type': 'column', 'subtype': 'stacked'})

            severity_col_letter = start_col_letter

            # Add "New" series with brand blue fill
            chart.add_series({
                'name': f"={sheet_name}!${chr(ord(start_col_letter)+1)}${start_row + 1}",
                'categories': f"={sheet_name}!${severity_col_letter}${start_row + 2}:${severity_col_letter}${start_row + 6}",
                'values': f"={sheet_name}!${chr(ord(start_col_letter)+1)}${start_row + 2}:${chr(ord(start_col_letter)+1)}${start_row + 6}",
                'fill': {'color': series_colors['New']},
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'font': {'size': 8, 'color': '#1F497D'},
                    'num_format': '[=0]"";General'  # Hide zero values
                }
            })

            # Add "Old" series with light blue fill
            chart.add_series({
                'name': f"={sheet_name}!${chr(ord(start_col_letter)+2)}${start_row + 1}",
                'categories': f"={sheet_name}!${severity_col_letter}${start_row + 2}:${severity_col_letter}${start_row + 6}",
                'values': f"={sheet_name}!${chr(ord(start_col_letter)+2)}${start_row + 2}:${chr(ord(start_col_letter)+2)}${start_row + 6}",
                'fill': {'color': series_colors['Old']},
                'data_labels': {
                    'value': True,
                    'position': 'outside_end',
                    'font': {'size': 8, 'color': '#1F497D'},
                    'num_format': '[=0]"";General'
                }
            })

            # Set chart title with matching font
            chart.set_title({
                'name': f"{bu} - {report_type.split('_')[0]} Report",
                'name_font': title_font
            })

            # Set X axis title and font
            chart.set_x_axis({
                'name': "Severity Levels",
                'name_font': axis_font,
                'num_font': axis_font
            })

            # Set Y axis title and font
            chart.set_y_axis({
                'name': "Count",
                'min': 0,
                'name_font': axis_font,
                'num_font': axis_font
            })

            # Set legend position and font
            chart.set_legend({
                'position': 'bottom',
                'font': legend_font
            })

            worksheet.insert_chart(chart_cell, chart)

    workbook.close()

if __name__ == "__main__":
    summary = process_reports()
    save_summary_to_excel(summary, "Vulnerability_Summary.xlsx")
