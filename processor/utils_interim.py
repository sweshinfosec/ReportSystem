import pandas as pd
import os
import socket
import struct
from django.conf import settings

def ip_to_numeric(ip):
    """Helper to convert IP strings to integers for accurate sorting."""
    try:
        return struct.unpack("!L", socket.inet_aton(str(ip).strip()))[0]
    except:
        return 4294967295

def generate_interim_report(filename_list, prefix="Pre_Remediation"):
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

    # 1. Rename 'Risk' to 'Severity'
    if 'Risk' in df.columns:
        df = df.rename(columns={'Risk': 'Severity'})
    
    if 'CVSS v3.0 Base Score' in df.columns:
        df['CVSS v3.0 Base Score'] = pd.to_numeric(df['CVSS v3.0 Base Score'], errors='coerce').fillna(0)

    # 2. Filter valid severities
    df = df.dropna(subset=['Severity', 'Host', 'Name'])
    df['Severity'] = df['Severity'].astype(str).str.strip()
    sev_order = ["Critical", "High", "Medium", "Low"]
    df = df[df['Severity'].isin(sev_order)]

    # 3. Multi-level Sort
    df['ip_sort_key'] = df['Host'].apply(ip_to_numeric)
    df['Severity'] = pd.Categorical(df['Severity'], categories=sev_order, ordered=True)
    df = df.sort_values(
        by=['Severity', 'Name', 'CVSS v3.0 Base Score', 'ip_sort_key'], 
        ascending=[True, True, False, True]
    )
    df = df.drop(columns=['ip_sort_key'])

    # 4. Final Columns (16 Columns)
    final_cols = [
        "Plugin ID", "CVE", "CVSS v3.0 Base Score", "Host", "Protocol", 
        "Port", "Name", "Synopsis", "Description", "Solution", 
        "Plugin Output", "Severity", "Vulnerability Status",
        "Opensource Comments", "Customer Comments", "Remarks"
    ]
    
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""

    df_interim = df[final_cols].copy().fillna("")

    # Save Files
    excel_name = f"{prefix}_Report.xlsx"
    csv_name = f"{prefix}_Report.csv"
    output_path = os.path.join(base_path, excel_name)
    df_interim.to_csv(os.path.join(base_path, csv_name), index=False)

    # 5. Pivot Tables
    pivot_host = pd.crosstab(df['Host'], df['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()
    pivot_findings = pd.crosstab(df['Name'], df['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()

    # --- SAVE EXCEL ---
    with pd.ExcelWriter(output_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        # Create sheets by writing the initial dataframes
        df_interim.to_excel(writer, index=False, sheet_name='InterimReport_VA')
        
        # Write a placeholder or the first pivot to 'Summary' to initialize it
        pivot_host.to_excel(writer, sheet_name='Summary', index=False, startcol=15, startrow=2)
        
        workbook = writer.book
        main_sheet = writer.sheets['InterimReport_VA']
        summary_sheet = writer.sheets['Summary']
        
        main_sheet.hide_gridlines(2)
        summary_sheet.hide_gridlines(2)
        
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D3D3D3', 'border': 1})
        body_fmt = workbook.add_format({'text_wrap': True, 'valign': 'top', 'border': 1})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#2B5797'})

        # Severity Coloring
        color_formats = {
            'Critical': workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}),
            'High':     workbook.add_format({'bg_color': '#FF9900', 'font_color': '#000000'}),
            'Medium':   workbook.add_format({'bg_color': '#FFFF00', 'font_color': '#000000'}),
            'Low':      workbook.add_format({'bg_color': '#00B050', 'font_color': '#FFFFFF'}),
        }

        # Main Sheet Formatting
        for i, col in enumerate(df_interim.columns):
            main_sheet.write(0, i, col, header_fmt)
            width = 45 if col in ["Description", "Solution", "Plugin Output", "Synopsis"] else 18
            main_sheet.set_column(i, i, width, body_fmt)

        sev_col_idx = df_interim.columns.get_loc("Severity")
        for sev_val, fmt in color_formats.items():
            main_sheet.conditional_format(1, sev_col_idx, len(df_interim), sev_col_idx, {
                'type': 'cell', 'criteria': 'equal to', 'value': f'"{sev_val}"', 'format': fmt
            })

        # --- Summary Tab Styling (Column P = Index 15) ---
        target_col = 15 
        summary_sheet.write(1, target_col, 'Vulnerability Count by Host', title_fmt)
        
        # Add Host Table
        summary_sheet.add_table(2, target_col, len(pivot_host) + 2, target_col + len(pivot_host.columns) - 1, {
            'data': pivot_host.values.tolist(),
            'columns': [{'header': c} for c in pivot_host.columns],
            'style': 'Table Style Medium 2'
        })

        # Add Findings Table
        start_row = len(pivot_host) + 6
        summary_sheet.write(start_row - 1, target_col, 'Unique Findings Summary', title_fmt)
        
        # We use to_excel to place the data, then add_table over it for styling
        pivot_findings.to_excel(writer, sheet_name='Summary', index=False, startcol=target_col, startrow=start_row)
        
        summary_sheet.add_table(start_row, target_col, start_row + len(pivot_findings), target_col + len(pivot_findings.columns) - 1, {
            'data': pivot_findings.values.tolist(),
            'columns': [{'header': c} for c in pivot_findings.columns],
            'style': 'Table Style Medium 2'
        })
        
        summary_sheet.set_column(target_col, target_col + 10, 22)

    return excel_name, csv_name