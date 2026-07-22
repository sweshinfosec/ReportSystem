import pandas as pd
import os
import socket
import struct
from django.conf import settings

def ip_to_int(ip):
    """Numeric IP sorting: 10.0.0.2 before 10.0.0.10."""
    try:
        return struct.unpack("!L", socket.inet_aton(str(ip).strip()))[0]
    except:
        return 0

def compare_pre_post_scans(pre_file_path, post_file_path, output_excel_name):
    base_path = settings.MEDIA_ROOT
    excel_path = os.path.join(base_path, output_excel_name)
    csv_output_name = output_excel_name.replace('.xlsx', '.csv')
    csv_path = os.path.join(base_path, csv_output_name)
    
    # 1. Load Data
    df_pre = pd.read_csv(pre_file_path)
    df_post = pd.read_csv(post_file_path)

    # 2. Pre-Processing
    def preprocess(df):
        if 'Risk' in df.columns:
            df.rename(columns={'Risk': 'Severity'}, inplace=True)
        df = df.loc[:, ~df.columns.duplicated()].copy()
        df['Severity'] = df['Severity'].astype(str).str.strip()
        df = df[~df['Severity'].str.contains('None|none|^$', na=True)].copy()
        df['uid'] = df['Name'].astype(str).str.strip() + \
                    df['Host'].astype(str).str.strip() + \
                    df['Port'].astype(str).str.strip()
        return df

    df_pre = preprocess(df_pre)
    df_post = preprocess(df_post)

    # 3. Status and Severity Mapping Logic
    severity_map = df_post.set_index('uid')['Severity'].to_dict()
    post_uids = set(df_post['uid'].unique())
    pre_uids = set(df_pre['uid'].unique())

    # 4. Process Round 1 Data (Existing findings)
    df_pre['Vulnerability Status'] = df_pre['uid'].apply(lambda x: "Open" if x in post_uids else "Fixed & Closed")
    df_pre['Severity'] = df_pre.apply(lambda row: severity_map.get(row['uid'], row['Severity']), axis=1)

    # Logic for Comments & Artefacts (R1 Findings)
    def get_r1_comments(row):
        if row['Vulnerability Status'] == "Open":
            opensource = "After the final scan Opensource team has confirmed that the issue still persists and the finding remains open in the reported environment."
            customer = "Infra team to add comments."
            artefact = "No evidence screenshot provided for validation."
        else:
            opensource = "After the final scan Opensource team has confirmed this vulnerability is not reported in the final round of VA scan hence marked as 'Fixed & Closed'."
            customer = "After the final scan Opensource team has confirmed this vulnerability is not reported in the final round of VA scan hence marked as 'Fixed & Closed’, No action needed."
            artefact = "The Nessus scan report and the interim report are attached and available in the APPENDIX section."
        return pd.Series([opensource, customer, artefact])

    df_pre[['Opensource Comments', 'Customer Comments', 'Artefacts']] = df_pre.apply(get_r1_comments, axis=1)

    # 5. Handle New Findings (Items in R2 but NOT in R1)
    new_findings = df_post[~df_post['uid'].isin(pre_uids)].copy()
    
    if not new_findings.empty:
        new_findings['Vulnerability Status'] = "Open"
        new_findings['Opensource Comments'] = "After the final scan Opensource team has confirmed this vulnerability newly reported in the final round of VA scan hence marking the status for reported servers as ‘Open’."
        new_findings['Customer Comments'] = "Infra team to add comments."
        new_findings['Artefacts'] = "No evidence screenshot provided for validation."
        final_df = pd.concat([df_pre, new_findings], ignore_index=True)
    else:
        final_df = df_pre.copy()

    # Fill blank manual columns
    if 'Remarks' not in final_df.columns:
        final_df['Remarks'] = ""

    # 6. Formatting/Trimming/Sorting Engine
    base_cols = ['Plugin ID', 'CVE', 'CVSS v3.0 Base Score', 'Host', 'Protocol', 'Port', 'Name', 'Synopsis', 'Description', 'Solution', 'Plugin Output', 'Severity']
    sev_order = ['Critical', 'High', 'Medium', 'Low']

    def clean_and_format(df, is_final=False):
        df['ip_int'] = df['Host'].apply(ip_to_int)
        # Sort by severity first (Critical, High, Medium, Low), same convention
        # as the Phase-1 interim report. Without this, the sheet was sorted
        # alphabetically by Name only, with Critical/High/Medium/Low findings
        # scattered throughout instead of grouped.
        sev_rank_map = {s: i for i, s in enumerate(sev_order)}
        df['sev_rank'] = df['Severity'].map(sev_rank_map).fillna(len(sev_order))
        df = df.sort_values(by=['sev_rank', 'Name', 'CVSS v3.0 Base Score', 'ip_int'], ascending=[True, True, False, True])
        df = df.drop(columns=['sev_rank'])

        current_cols = base_cols.copy()
        if is_final:
            current_cols.append('Vulnerability Status')
            # Column Sequence including Artefacts
            current_cols.extend(['Opensource Comments', 'Customer Comments', 'Remarks', 'Artefacts'])
            
        return df[[c for c in current_cols if c in df.columns]].copy()

    trimmed_r1 = clean_and_format(df_pre, is_final=False)
    trimmed_r2 = clean_and_format(df_post, is_final=False)
    trimmed_final = clean_and_format(final_df, is_final=True)

    # 7. Save Files
    trimmed_final.to_csv(csv_path, index=False)

    with pd.ExcelWriter(excel_path, engine='xlsxwriter', engine_kwargs={'options': {'nan_inf_to_errors': True}}) as writer:
        trimmed_r1.to_excel(writer, sheet_name='InterimReport_VA_R1', index=False)
        trimmed_r2.to_excel(writer, sheet_name='InterimReport_VA_R2', index=False)
        trimmed_final.to_excel(writer, sheet_name='InterimReport_VA', index=False)
        
        workbook = writer.book
        main_sheet = writer.sheets['InterimReport_VA']
        summary_sheet = workbook.add_worksheet('Summary')
        summary_sheet.hide_gridlines(2)
        
        # Formats
        header_fmt = workbook.add_format({'bold': True, 'bg_color': '#D7E4BC', 'border': 1})
        title_fmt = workbook.add_format({'bold': True, 'font_size': 14, 'font_color': '#2B5797'})
        color_fmts = {
            'Critical': workbook.add_format({'bg_color': '#8B0000', 'font_color': '#FFFFFF'}),
            'High':     workbook.add_format({'bg_color': '#FF0000', 'font_color': '#FFFFFF'}),
            'Medium':   workbook.add_format({'bg_color': '#FFC000'}),
            'Low':      workbook.add_format({'bg_color': '#FFFF00'})
        }

        # Main Tab Headers and Coloring
        for i, col in enumerate(trimmed_final.columns):
            main_sheet.write(0, i, col, header_fmt)
        
        sev_idx = trimmed_final.columns.get_loc("Severity")
        for val, fmt in color_fmts.items():
            main_sheet.conditional_format(1, sev_idx, len(trimmed_final), sev_idx, 
                {'type': 'cell', 'criteria': 'equal to', 'value': f'"{val}"', 'format': fmt})

        # Summary Pivots
        pivot_host = pd.crosstab(trimmed_final['Host'], trimmed_final['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()
        pivot_findings = pd.crosstab(trimmed_final['Name'], trimmed_final['Severity']).reindex(columns=sev_order, fill_value=0).reset_index()

        target_col = 15
        summary_sheet.write(1, target_col, 'Vulnerability Count by Host', title_fmt)
        summary_sheet.add_table(2, target_col, len(pivot_host) + 2, target_col + len(pivot_host.columns) - 1, {
            'data': pivot_host.values.tolist(), 'columns': [{'header': c} for c in pivot_host.columns], 'style': 'Table Style Medium 2'
        })

        start_row = len(pivot_host) + 6
        summary_sheet.write(start_row - 1, target_col, 'Unique Findings Summary', title_fmt)
        summary_sheet.add_table(start_row, target_col, start_row + len(pivot_findings), target_col + len(pivot_findings.columns) - 1, {
            'data': pivot_findings.values.tolist(), 'columns': [{'header': c} for c in pivot_findings.columns], 'style': 'Table Style Medium 2'
        })
        
        main_sheet.set_column('A:Z', 18)
        summary_sheet.set_column(target_col, target_col + 10, 25)

    return output_excel_name, csv_output_name