# Streamlit-compatible version of the warehouse packing report with improved color parsing

import streamlit as st
import pandas as pd
import re
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from io import BytesIO

st.title("📦 VCC Warehouse Packing Group Generator")

uploaded_file = st.file_uploader("Upload your POInput file (CSV or Excel)", type=["xlsx", "csv"])

if uploaded_file is not None:
    
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

        # Infant sizes
        ws.cell(row=current_row, column=1, value="ColorStyle")
        for i, size in enumerate(infant_sizes, start=2):
            ws.cell(row=current_row, column=i, value=size)
        ws.cell(row=current_row, column=i+1, value="Total")
        current_row += 1

        for _, row in group_data.iterrows():
            if row['Infant Total'] > 0:
                ws.cell(row=current_row, column=1, value=row['ColorStyle'])
                for i, size in enumerate(infant_sizes, start=2):
                    ws.cell(row=current_row, column=i, value=row.get(size, 0))
                ws.cell(row=current_row, column=i+1, value=row['Infant Total'])
                current_row += 1

        current_row += 1

        # Toddler sizes
        ws.cell(row=current_row, column=1, value="ColorStyle")
        for i, size in enumerate(toddler_sizes, start=2):
            ws.cell(row=current_row, column=i, value=size)
        ws.cell(row=current_row, column=i+1, value="Total")
        current_row += 1

        for _, row in group_data.iterrows():
            if row['Toddler Total'] > 0:
                ws.cell(row=current_row, column=1, value=row['ColorStyle'])
                for i, size in enumerate(toddler_sizes, start=2):
                    ws.cell(row=current_row, column=i, value=row.get(size, 0))
                ws.cell(row=current_row, column=i+1, value=row['Toddler Total'])
                current_row += 1

        # Associated POs
        current_row += 1
        ws.cell(row=current_row, column=1, value="Associated POs:")
        po_list_split = po_list.split(', ')
        for i, po in enumerate(po_list_split):
            ws.cell(row=current_row + i, column=2, value=po)
        current_row += len(po_list_split)
        current_row += 2

    output = BytesIO()
    wb.save(output)
    output.seek(0)

    st.success("✅ Report generated!")
    st.download_button("📥 Download Packing Report", output, file_name="Packing_Group_Report.xlsx")
