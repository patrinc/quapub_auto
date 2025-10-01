import os
import sys
import pandas as pd

# Add current directory (root) to sys.path so config.py can be imported
current_dir = os.path.dirname(os.path.abspath(__file__))
if current_dir not in sys.path:
    sys.path.insert(0, current_dir)

import config

YEAR = config.year
CURRENT_MONTH = config.current_month
CURRENT_MONTH_NUM = config.current_month_num

# Locally define these lists since they are not in config.py
COUNTRIES = ['Brazil', 'CSA', 'India', 'Mexico', 'SFTL']
REPORT_TYPES = ['Network', 'Server']

SUBJECT = f"Approval Request: Qualys Monthly Scan Reports for {CURRENT_MONTH} {YEAR}"

FINAL_REPORTS_DIR = rf"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning\{YEAR[2:]}_{CURRENT_MONTH_NUM} Monthly Scans\01_with_Status_and_Overview\For Approval"
SUMMARY_DIR = rf"C:\Users\example.user1\OneDrive - exampledomain\Assignments\VAPT\Infra Scanning\{YEAR[2:]}_{CURRENT_MONTH_NUM} Monthly Scans\03_Summary Requested"
SUMMARY_FILE = "Vulnerability_Summary.xlsx"
SUMMARY_PATH = os.path.join(SUMMARY_DIR, SUMMARY_FILE)
SHAREPOINT_FOLDER_LINK = "https://exampledomaindigital.sharepoint.com/:f:/r/sites/GSRCGlobal/Shared%20Documents/General/CIC/Vulnerability%20Mgmt/Qualys%20Vulnerability%20Status%20Records?csf=1&web=1&e=Ac1Gpy"

def extract_full_summary_html(summary_path):
    try:
        xls = pd.ExcelFile(summary_path, engine='openpyxl')

        network_dfs = []
        server_dfs = []

        for sheet_name in xls.sheet_names:
            try:
                # Read Network Report section (rows 2-6)
                df_network = pd.read_excel(
                    xls, sheet_name=sheet_name,
                    usecols=["Severity", "New", "Old"],
                    skiprows=1, nrows=5
                )
                df_network["Severity"] = df_network["Severity"].astype(str).str.strip()
                df_network.set_index("Severity", inplace=True)
                # Add BU name for columns
                df_network.columns = pd.MultiIndex.from_product([[sheet_name], df_network.columns])
                network_dfs.append(df_network)

                # Read Server Report section (rows 20-24)
                df_server = pd.read_excel(
                    xls, sheet_name=sheet_name,
                    usecols=["Severity", "New", "Old"],
                    skiprows=19, nrows=5
                )
                df_server["Severity"] = df_server["Severity"].astype(str).str.strip()
                df_server.set_index("Severity", inplace=True)
                df_server.columns = pd.MultiIndex.from_product([[sheet_name], df_server.columns])
                server_dfs.append(df_server)

            except Exception as e:
                print(f"Error reading sheet {sheet_name}: {e}")
                continue

        # Combine all BUs horizontally for each report type
        network_all = pd.concat(network_dfs, axis=1)
        server_all = pd.concat(server_dfs, axis=1)

        # Now add top-level multiindex for report type
        network_all.columns = pd.MultiIndex.from_tuples(
            [("Network", bu, metric) for bu, metric in network_all.columns],
            names=["Report Type", "BU", "Metric"]
        )
        server_all.columns = pd.MultiIndex.from_tuples(
            [("Server", bu, metric) for bu, metric in server_all.columns],
            names=["Report Type", "BU", "Metric"]
        )

        # Concatenate Network and Server side by side
        combined_df = pd.concat([network_all, server_all], axis=1)

        # Sort columns by Report Type, then BU, then Metric
        combined_df = combined_df.sort_index(axis=1, level=[0,1,2])

        # Style the combined dataframe for HTML email
        styled = combined_df.style.set_table_styles([
            {'selector': 'th', 'props': [
                ('background-color', '#D9E1F2'),
                ('color', '#1F497D'),
                ('border', '1px solid #4F81BD')]},
            {'selector': 'td', 'props': [
                ('border', '1px solid #4F81BD'),
                ('text-align', 'center')]},
        ]).set_properties(**{'font-family': 'Calibri', 'font-size': '11pt'})

        html_table = styled.to_html()
        print("Summary table extracted and styled successfully.")
        return html_table

    except Exception as e:
        print(f"Error extracting summary table: {e}")
        return "<p><em>Summary table not available.</em></p>"

def create_email_body(summary_html, charts_note, sharepoint_folder_link):
    html_body = f"""
    <div style="font-family:Calibri; font-size:11pt; color:#1F497D; max-width: 900px;">
      <p>Hi Ashok,</p>
      <p>Please find attached the final Qualys monthly scan reports for <strong>{CURRENT_MONTH} {YEAR}</strong>.</p>
      <p><strong>Summary Table:</strong></p>
      {summary_html}
      {charts_note}
      <p>All vulnerability records with their status values are available in the SharePoint folder below:</p>
      <p><a href="{sharepoint_folder_link}">SharePoint: Qualys Vulnerability Status Records</a></p>
      <p>Kindly review the attached reports and provide your approval to proceed with sharing them with the Tech IOs.</p>
      <p>Let me know if you have any questions or require further details.</p>
    </div>
    """
    return html_body

def attach_reports(mail, reports_dir):
    attached_files = 0
    if not os.path.exists(reports_dir):
        print(f"Final reports directory does not exist: {reports_dir}")
        return attached_files

    for country in COUNTRIES:
        for report_type in REPORT_TYPES:
            expected_filename = f"{country}_{YEAR}_{CURRENT_MONTH_NUM}_Final_Report_{report_type}_Report.xlsx"
            filepath = os.path.join(reports_dir, expected_filename)
            if os.path.isfile(filepath):
                mail.Attachments.Add(filepath)
                attached_files += 1
            else:
                print(f"Expected report file not found: {filepath}")

    if attached_files == 0:
        print(f"No final report Excel files attached from {reports_dir}.")
    else:
        print(f"Attached {attached_files} final report files.")
    return attached_files

def attach_summary(mail, summary_path):
    if os.path.exists(summary_path):
        mail.Attachments.Add(summary_path)
        print(f"Attached summary file: {summary_path}")
        return 1
    else:
        print(f"Summary file not found: {summary_path}")
        return 0

def send_approval_email():
    summary_html = extract_full_summary_html(SUMMARY_PATH)
    try:
        import win32com.client
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
    except Exception as e:
        print(f"Error initializing Outlook: {e}")
        return
    mail.Subject = SUBJECT
    mail.To = "example.user2@exampledomain.com"
    mail.Cc = "example.user3@exampledomain.com; example.user4@exampledomain.com"
    mail.Bcc = "example.user1@exampledomain.com"
    mail.Display()
    default_signature = mail.HTMLBody
    charts_note = f"""
    <p>All summarized charts by Business Unit are available in the attached Excel file <strong>{SUMMARY_FILE}</strong>, which contains BU-wise sheets.</p>
    """
    mail.HTMLBody = create_email_body(summary_html, charts_note, SHAREPOINT_FOLDER_LINK) + default_signature
    attached_files = attach_reports(mail, FINAL_REPORTS_DIR)
    attached_files += attach_summary(mail, SUMMARY_PATH)
    print(f"Email created and displayed successfully with {attached_files} attachments.")
    print("Please review and send manually.")

if __name__ == "__main__":
    send_approval_email()
