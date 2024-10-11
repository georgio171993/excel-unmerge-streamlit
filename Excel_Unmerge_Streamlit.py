import openpyxl
import streamlit as st
from datetime import datetime

# Function to get the value of a cell, accounting for merged cells
def get_merged_cell_value(ws, row, col):
    """Return the value of a cell, considering merged cells."""
    cell = ws.cell(row=row, column=col)
    if cell.value is None:  # If the cell is part of a merged cell, find the top-left cell
        for merge in ws.merged_cells.ranges:
            if (merge.min_row <= row <= merge.max_row) and (merge.min_col <= col <= merge.max_col):
                # Return the value from the top-left cell of the merged range
                return ws.cell(merge.min_row, merge.min_col).value
    return cell.value

# Function to process the Excel file
def process_excel(file):
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    # 3. Replace '1_4_10' value with '1_1_9' if '1_4_10' is blank and '1_4_8' is not blank
    st.write("Applying the condition for blank values in '1_4_10' and '1_4_8'...")
    
    for row in range(2, ws.max_row + 1):  # Assuming the first row is the header
        cell_1_4_10 = ws.cell(row=row, column=52)  # '1_4_10' is the 52nd column
        cell_1_4_8 = ws.cell(row=row, column=50)   # '1_4_8' is the 50th column (corrected)
        
        # Fetch the value from '1_1_9', considering merged cells
        cell_1_1_9_value = get_merged_cell_value(ws, row, 9)  # '1_1_9' is the 9th column
        
        # Check if '1_4_10' is blank and '1_4_8' is not blank
        if cell_1_4_10.value in [None, ""] and cell_1_4_8.value not in [None, ""]:
            cell_1_4_10.value = cell_1_1_9_value  # Replace '1_4_10' with the value from '1_1_9'

    # Create output file name with the current date and time
    current_time = datetime.now().strftime('%d %b %Y %H:%M')
    output_file = f"Questionnaire Answers - {current_time}.xlsx"
    
    wb.save(output_file)
    return output_file

# Streamlit app
st.title("Excel Unmerge and Update Tool")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    st.write("Processing...")
    output_file = process_excel(uploaded_file)
    
    with open(output_file, "rb") as f:
        st.download_button("Download the processed file", f, file_name=output_file)
