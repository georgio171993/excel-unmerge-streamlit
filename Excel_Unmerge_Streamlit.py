
import openpyxl
import streamlit as st
from datetime import datetime
import os

# Function to process the Excel file
def process_excel(file):
    wb = openpyxl.load_workbook(file)
    ws = wb.active
    
    for merge in list(ws.merged_cells.ranges):
        min_row, min_col, max_row, max_col = merge.min_row, merge.min_col, merge.max_row, merge.max_col
        
        ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        
        first_cell_value = ws.cell(row=min_row, column=min_col).value
        
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if row == min_row and col == min_col:
                    continue
                ws.cell(row=row, column=col).value = None

    today_date = datetime.today().strftime('%Y-%m-%d')
    output_file = f"unmerged_output_{today_date}.xlsx"
    
    wb.save(output_file)
    return output_file

# Streamlit app
st.title("Excel Unmerge Tool")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("Processing...")
    output_file = process_excel(uploaded_file)
    
    with open(output_file, "rb") as f:
        st.download_button("Download the processed file", f, file_name=output_file)
