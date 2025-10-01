import os
import sys
import win32com.client as win32

# Add parent directory (where config.py is located) to sys.path to allow import
current_dir = os.path.dirname(os.path.abspath(__file__))
parent_dir = os.path.abspath(os.path.join(current_dir, '..'))
if parent_dir not in sys.path:
    sys.path.insert(0, parent_dir)

import config  # Access year, current_month_num from config.py

def export_excel_charts_to_png():
    base_path = r"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning"
    folder_name = f"{config.year[-2:]}_{config.current_month_num} Monthly Scans"
    summary_subfolder = "03_Summary Requested"
    summary_file = os.path.join(base_path, folder_name, summary_subfolder, "Vulnerability_Summary.xlsx")
    charts_output_folder = os.path.join(base_path, folder_name, summary_subfolder, "Exported_Charts")

    if not os.path.exists(charts_output_folder):
        os.makedirs(charts_output_folder)

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(summary_file)
    try:
        for sheet in wb.Sheets:
            chart_objects = sheet.ChartObjects()
            for i in range(1, chart_objects.Count + 1):
                chart = chart_objects.Item(i).Chart
                suffix = "Network" if i == 1 else "Server" if i == 2 else f"Chart{i}"
                filename = f"{sheet.Name}_{suffix}.png"
                full_path = os.path.join(charts_output_folder, filename)
                chart.Export(full_path)
                print(f"Exported chart from sheet '{sheet.Name}' as {full_path}")
    finally:
        wb.Close(SaveChanges=False)
        excel.Quit()

if __name__ == "__main__":
    export_excel_charts_to_png()
