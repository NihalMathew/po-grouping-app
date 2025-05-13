#app.py

import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment

st.title("📦 Warehouse Packing Group Generator")

uploaded_file = st.file_uploader("Upload your POInput.xlsx file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    def extract_color(desc):
        if pd.isna(desc):
            return ''
        parts = desc.split('_')
        for part in reversed(parts):
            clean = part.strip()
            if clean.isalpha() and len(clean) >= 3:
                return clean.upper()
        if ',' in desc:
            for segment in desc.split(','):
                word = segment.strip()
                if word.isalpha() and len(word) >= 3:
                    return word.upper()
        return 'UNKNOWN'

    def extract_style_digits(style):
        if pd.isna(style):
            return ''
        match = re.search(r'(\d+)$', style)
        return match.group(1) if match else ''

    df['Color'] = df['Material Description'].apply(extract_color)
    df['StyleDigits'] = df['Style Code'].apply(extract_style_digits)
    df['ColorStyle'] = df['Color'] + ' - ' + df['StyleDigits']

    pivot = df.pivot_table(
        index=['PO Number', 'ColorStyle'],
        columns='Size',
        values='Article Qty',
        aggfunc='sum',
        fill_value=0
    )

    pivot['Total'] = pivot.sum(axis=1)
    pivot.reset_index(inplace=True)

    size_cols = ['6-12M', '12-18M', '18-24M']
    for col in size_cols:
        if col not in pivot.columns:
            pivot[col] = 0
    pivot = pivot[['PO Number', 'ColorStyle'] + size_cols + ['Total']]

    po_group_map = defaultdict(list)
    for po_number, group in pivot.groupby('PO Number'):
        signature = tuple(sorted(
            tuple([row['ColorStyle']] + [int(row[col]) for col in size_cols])
            for _, row in group.iterrows()
        ))
        po_group_map[signature].append(po_number)

    grouped_rows = []
    for idx, (sig, po_list) in enumerate(sorted(po_group_map.items()), start=1):
        po_count = len(set(po_list))
        for entry in sig:
            color_style, s1, s2, s3 = entry
            total = s1 + s2 + s3
            grouped_rows.append({
                'Group ID': f'Group {idx}',
                'ColorStyle': color_style,
                '6-12M': s1,
                '12-18M': s2,
                '18-24M': s3,
                'Total': total,
                'POs': ', '.join(map(str, sorted(set(po_list)))),
                'PO Count': po_count
            })

    grouped_df = pd.DataFrame(grouped_rows)
    grouped_df_sorted = grouped_df.sort_values(by='Group ID')
    group_ids = sorted(
        grouped_df_sorted['Group ID'].unique(),
        key=lambda g: int(g.replace('Group ', ''))
    )

    # Generate Excel in-memory
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing Groups"
    current_row = 1

    for group_id in group_ids:
        group_data = grouped_df_sorted[grouped_df_sorted['Group ID'] == group_id]
        po_list = group_data['POs'].iloc[0]
        po_count = group_data['PO Count'].iloc[0]

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        header_cell = ws.cell(row=current_row, column=1, value=f"{group_id} (PO Count: {po_count})")
        header_cell.font = Font(bold=True, size=12)
        header_cell.alignment = Alignment(horizontal='center')
        current_row += 1

        headers = ['ColorStyle', '6-12M', '12-18M', '18-24M', 'Total']
        for col_idx, header in enumerate(headers, start=1):
            cell = ws.cell(row=current_row, column=col_idx, value=header)
            cell.font = Font(bold=True)
        current_row += 1

        for _, row in group_data.iterrows():
            ws.cell(row=current_row, column=1, value=row['ColorStyle'])
            ws.cell(row=current_row, column=2, value=row['6-12M'])
            ws.cell(row=current_row, column=3, value=row['12-18M'])
            ws.cell(row=current_row, column=4, value=row['18-24M'])
            ws.cell(row=current_row, column=5, value=row['Total'])
            current_row += 1

        current_row += 1
        ws.cell(row=current_row, column=1, value="Associated POs:")
        po_lines = '\n'.join(po_list.split(', '))
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=6)
        po_cell = ws.cell(row=current_row, column=2, value=po_lines)
        po_cell.alignment = Alignment(wrapText=True, vertical='top')
        current_row += 2

    # Save to BytesIO
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Report generated!")
    st.download_button("📥 Download Packing Report", output, file_name="Packing_Group_Report.xlsx")
