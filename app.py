# Final updated script for warehouse packing report with two size bands and grouped POs

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from collections import defaultdict
import pandas as pd
import re
from io import BytesIO
import streamlit as st

st.title("📦 Warehouse Packing Group Generator")

uploaded_file = st.file_uploader("Upload your POInput.xlsx file", type="xlsx")

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    required_cols = {'PO Number', 'Material Description', 'Style Code', 'Size', 'Article Qty'}
    if not required_cols.issubset(set(df.columns)):
        st.error(f"❌ Uploaded file is missing required columns. Please include: {', '.join(required_cols)}")
        st.stop()

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

    infant_sizes = ['6-12M', '12-18M', '18-24M']
    toddler_sizes = ['2-3Y', '3-4Y', '5-6Y', '7-8Y']
    all_sizes = infant_sizes + toddler_sizes

    pivot = df.pivot_table(
        index=['PO Number', 'ColorStyle'],
        columns='Size',
        values='Article Qty',
        aggfunc='sum',
        fill_value=0
    )

    for col in all_sizes:
        if col not in pivot.columns:
            pivot[col] = 0

    pivot['Infant Total'] = pivot[infant_sizes].sum(axis=1)
    pivot['Toddler Total'] = pivot[toddler_sizes].sum(axis=1)
    pivot.reset_index(inplace=True)

    po_group_map = defaultdict(list)
    for po_number, group in pivot.groupby('PO Number'):
        signature = tuple(sorted(
            tuple([row['ColorStyle']] + [int(row.get(col, 0)) for col in all_sizes])
            for _, row in group.iterrows()
        ))
        po_group_map[signature].append(po_number)

    grouped_rows = []
    for idx, (sig, po_list) in enumerate(sorted(po_group_map.items()), start=1):
        po_count = len(set(po_list))
        for entry in sig:
            color_style = entry[0]
            size_vals = entry[1:]
            infant_values = size_vals[:len(infant_sizes)]
            toddler_values = size_vals[len(infant_sizes):]
            grouped_rows.append({
                'Group ID': f'Group {idx}',
                'ColorStyle': color_style,
                **{infant_sizes[i]: infant_values[i] for i in range(len(infant_sizes))},
                **{toddler_sizes[i]: toddler_values[i] for i in range(len(toddler_sizes))},
                'Infant Total': sum(infant_values),
                'Toddler Total': sum(toddler_values),
                'POs': ', '.join(map(str, sorted(set(po_list)))),
                'PO Count': po_count
            })

    grouped_df = pd.DataFrame(grouped_rows)
    grouped_df_sorted = grouped_df.sort_values(by='Group ID')
    group_ids = sorted(grouped_df_sorted['Group ID'].unique(), key=lambda g: int(g.replace('Group ', '')))

    wb = Workbook()
    ws = wb.active
    ws.title = "Packing Groups"
    current_row = 1

    for group_id in group_ids:
        group_data = grouped_df_sorted[grouped_df_sorted['Group ID'] == group_id]
        po_list = group_data['POs'].iloc[0]
        po_count = group_data['PO Count'].iloc[0]

        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=8)
        header_cell = ws.cell(row=current_row, column=1, value=f"{group_id} (PO Count: {po_count})")
        header_cell.font = Font(bold=True, size=12)
        header_cell.alignment = Alignment(horizontal='center')
        current_row += 1

        # Infant size table
        ws.cell(row=current_row, column=1, value="ColorStyle")
        for i, size in enumerate(infant_sizes, start=2):
            ws.cell(row=current_row, column=i, value=size)
        ws.cell(row=current_row, column=i+1, value="Total")
        current_row += 1

        for _, row in group_data.iterrows():
            ws.cell(row=current_row, column=1, value=row['ColorStyle'])
            for i, size in enumerate(infant_sizes, start=2):
                ws.cell(row=current_row, column=i, value=row.get(size, 0))
            ws.cell(row=current_row, column=i+1, value=row['Infant Total'])
            current_row += 1

        current_row += 1

        # Toddler size table
        ws.cell(row=current_row, column=1, value="ColorStyle")
        for i, size in enumerate(toddler_sizes, start=2):
            ws.cell(row=current_row, column=i, value=size)
        ws.cell(row=current_row, column=i+1, value="Total")
        current_row += 1

        for _, row in group_data.iterrows():
            ws.cell(row=current_row, column=1, value=row['ColorStyle'])
            for i, size in enumerate(toddler_sizes, start=2):
                ws.cell(row=current_row, column=i, value=row.get(size, 0))
            ws.cell(row=current_row, column=i+1, value=row['Toddler Total'])
            current_row += 1

        current_row += 1
        ws.cell(row=current_row, column=1, value="Associated POs:")
        po_lines = '\n'.join(po_list.split(', '))
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=8)
        po_cell = ws.cell(row=current_row, column=2, value=po_lines)
        po_cell.alignment = Alignment(wrapText=True, vertical='top')
        current_row += 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Report generated!")
    st.download_button("📥 Download Packing Report", output, file_name="Packing_Group_Report.xlsx")
