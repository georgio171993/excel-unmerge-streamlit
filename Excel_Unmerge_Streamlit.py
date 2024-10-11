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

# Function to set the value of a cell, accounting for merged cells
def set_merged_cell_value(ws, row, col, value):
    """Set the value of a cell, handling merged cells."""
    # Check if the cell is part of a merged range
    for merge in ws.merged_cells.ranges:
        if (merge.min_row <= row <= merge.max_row) and (merge.min_col <= col <= merge.max_col):
            # Set the value only in the top-left cell of the merged range
            ws.cell(merge.min_row, merge.min_col).value = value
            return
    # If the cell is not part of a merged range, set the value directly
    ws.cell(row=row, column=col).value = value

# Function to process the Excel file
def process_excel(file):
    wb = openpyxl.load_workbook(file)
    ws = wb.active

    # 3. Replace '1_4_10' value with '1_1_9' if '1_4_10' is blank and '1_4_4' is not blank
    st.write("Applying the condition for blank values in '1_4_10' and '1_4_4'...")
    
    for row in range(2, ws.max_row + 1):  # Assuming the first row is the header
        cell_1_4_10 = ws.cell(row=row, column=52)  # '1_4_10' is the 52nd column
        cell_1_4_4 = ws.cell(row=row, column=46)   # '1_4_4' is the 46th column
        
        # Fetch the value from '1_1_9', considering merged cells
        cell_1_1_9_value = get_merged_cell_value(ws, row, 9)  # '1_1_9' is the 9th column
        
        # Check if '1_4_10' is blank and '1_4_4' is not blank
        if cell_1_4_10.value in [None, ""] and cell_1_4_4.value not in [None, ""]:
            set_merged_cell_value(ws, row, 52, cell_1_1_9_value)  # Replace '1_4_10' with the value from '1_1_9'
            
    # 1. Unmerge cells and handle the unmerged values (existing logic)
    st.write("Unmerging other cells and handling values...")
    
    for merge in list(ws.merged_cells.ranges):
        min_row, min_col, max_row, max_col = merge.min_row, merge.min_col, merge.max_row, merge.max_col
        
        ws.unmerge_cells(start_row=min_row, start_column=min_col, end_row=max_row, end_column=max_col)
        
        first_cell_value = ws.cell(row=min_row, column=min_col).value
        
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                if row == min_row and col == min_col:
                    continue
                ws.cell(row=row, column=col).value = None
    
    # 2. Condition to replace 'N/A' with blanks
    st.write("Replacing 'N/A' with blanks...")
    
    for row in ws.iter_rows():
        for cell in row:
            if cell.value == 'N/A':
                cell.value = ''

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
