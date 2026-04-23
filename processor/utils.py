import pandas as pd
import os
import socket
import struct
from django.conf import settings

def ip_to_numeric(ip):
    try:
        return struct.unpack("!L", socket.inet_aton(str(ip).strip()))[0]
    except:
        return 4294967295

def generate_interim_report(filename_list):
    base_path = settings.MEDIA_ROOT
    all_dfs = []

    for filename in filename_list:
        file_path = os.path.join(base_path, filename)
        try:
            temp = pd.read_csv(file_path, low_memory=False, on_bad_lines='skip')
            all_dfs.append(temp)
        except:
            continue
    
    if not all_dfs: return None, None
    df = pd.concat(all_dfs, ignore_index=True).drop_duplicates()

    # Rename & Clean
    if 'Risk' in df.columns:
        df = df.rename(columns={'Risk': 'Severity'})
    
    df = df.dropna(subset=['Severity', 'Host', 'Name'])
    df['Severity'] = df['Severity'].astype(str).str.strip()
    df = df[df['Severity'].str.lower() != 'none']
    sev_order = ["Critical", "High", "Medium", "Low"]
    df = df[df['Severity'].isin(sev_order)]

    # Multi-level Sort
    df['ip_sort_key'] = df['Host'].apply(ip_to_numeric)
    df['Severity'] = pd.Categorical(df['Severity'], categories=sev_order, ordered=True)
    df = df.sort_values(by=['Severity', 'Name', 'CVSS v3.0 Base Score', 'ip_sort_key'], ascending=[True, True, False, True])
    df = df.drop(columns=['ip_sort_key'])

    # Final Columns for both Excel and CSV
    desired_cols = ["Plugin ID", "CVE", "CVSS v3.0 Base Score", "Host", "Protocol", "Port", "Name", "Synopsis", "Description", "Solution", "Plugin Output", "Severity"]
    existing = [c for c in desired_cols if c in df.columns]
    df_interim = df[existing].copy()
    for col in ["Opensource Comments", "Customer Comments", "Remarks"]:
        df_interim[col] = ""

    # --- SAVE CSV VERSION ---
    csv_name = "Merged_Scan_Report.csv"
    df_interim.to_csv(os.path.join(base_path, csv_name), index=False)

    # --- SAVE EXCEL VERSION ---
    excel_name = "Merged_Scan_Report.xlsx"
    output_path = os.path.join(base_path, excel_name)
    
    pivot_host = pd.crosstab(df['Host'], df['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()
    pivot_findings = pd.crosstab(df['Name'], df['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        df_interim.to_excel(writer, index=False, sheet_name='InterimReport_VA')
        pivot_host.to_excel(writer, sheet_name='Summary', startcol=14, startrow=1, index=False)
        pivot_findings.to_excel(writer, sheet_name='Summary', startcol=14, startrow=len(pivot_host) + 6, index=False)

        workbook = writer.book
        main_sheet = writer.sheets['InterimReport_VA']
        summary_sheet = writer.sheets['Summary']
        
        # Style Summary Tab
        summary_sheet.add_table(1, 14, len(pivot_host) + 1, 14 + len(sev_order), {
            'data': pivot_host.values.tolist(),
            'columns': [{'header': c} for c in pivot_host.columns],
            'style': 'Table Style Medium 2'
        })
        
        # Style Main Tab
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        for i, col in enumerate(df_interim.columns):
            main_sheet.write(0, i, col, header_fmt)
            width = 45 if col in ["Description", "Solution", "Plugin Output"] else 18
            main_sheet.set_column(i, i, width, workbook.add_format({'text_wrap': True, 'valign': 'top'}))

    return excel_name, csv_name
