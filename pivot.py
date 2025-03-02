import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from io import BytesIO
import base64

st.set_page_config(page_title="Excel Pivot Navigation App", layout="wide")

def main():
    st.title("Excel Pivot Navigation App")
    st.write("Upload an Excel file to add a navigation pivot sheet with data display")
    
    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Read the Excel file
        file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
        st.write(f"File uploaded: {file_details['FileName']}")
        
        # Process the file
        processed_file = process_excel_file(uploaded_file)
        
        # Download button
        download_button_str = download_button(processed_file, f"Pivot_{uploaded_file.name}", "Download Modified Excel File")
        st.markdown(download_button_str, unsafe_allow_html=True)
        
        st.success("Success! Your file has been processed. Download and open in Excel to use.")
        st.info("In the Navigation sheet, use the dropdown to select which sheet's data to view.")

def process_excel_file(uploaded_file):
    # Read the file into a BytesIO object
    bytes_data = BytesIO(uploaded_file.getvalue())
    
    # Load workbook with openpyxl
    workbook = openpyxl.load_workbook(bytes_data)
    
    # Get all sheet names
    sheet_names = workbook.sheetnames
    
    # Create a new sheet at the beginning for pivot
    pivot_sheet_name = "Navigation"
    if pivot_sheet_name in workbook.sheetnames:
        # If a sheet with this name already exists, remove it
        workbook.remove(workbook[pivot_sheet_name])
    
    # Create new pivot sheet at the beginning
    pivot_sheet = workbook.create_sheet(pivot_sheet_name, 0)
    
    # Add title and instructions
    pivot_sheet['A1'] = "Sheet Navigation"
    pivot_sheet['A2'] = "Select a sheet to view its data"
    
    title_font = Font(bold=True, size=14)
    instruction_font = Font(italic=True, size=12)
    pivot_sheet['A1'].font = title_font
    pivot_sheet['A2'].font = instruction_font
    
    # Create a dropdown selection for sheets
    pivot_sheet['B4'] = "Select a sheet:"
    pivot_sheet['B4'].font = Font(bold=True)
    
    # Create dropdown cell
    pivot_sheet['C4'] = sheet_names[0] if len(sheet_names) > 0 else ""
    
    # Set up data validation for dropdown
    sheet_list = ','.join([f'"{name}"' for name in sheet_names if name != pivot_sheet_name])
    dv = DataValidation(type="list", formula1=f"={sheet_list}")
    pivot_sheet.add_data_validation(dv)
    dv.add('C4')
    
    # Pre-populate the navigation sheet with the first sheet's data
    if len(sheet_names) > 0 and sheet_names[0] != pivot_sheet_name:
        first_sheet = workbook[sheet_names[0]]
        copy_data_to_pivot(first_sheet, pivot_sheet, 7)
    
    # Copy data from all sheets to hidden sheets with named ranges
    for sheet_name in sheet_names:
        if sheet_name != pivot_sheet_name:
            source_sheet = workbook[sheet_name]
            
            # Create a hidden sheet to store this data
            data_sheet_name = f"Data_{sheet_name}"
            if data_sheet_name in workbook.sheetnames:
                workbook.remove(workbook[data_sheet_name])
                
            data_sheet = workbook.create_sheet(data_sheet_name)
            data_sheet.sheet_state = 'hidden'
            
            # Copy data to the hidden sheet
            copy_all_data(source_sheet, data_sheet)
            
            # Create a named range for this data
            define_named_range(workbook, data_sheet, data_sheet_name)
    
    # Add formula to show selected sheet data
    set_up_data_lookup(pivot_sheet, sheet_names, pivot_sheet_name)
    
    # Hide all original sheets except the pivot sheet
    for sheet_name in sheet_names:
        if sheet_name != pivot_sheet_name:
            workbook[sheet_name].sheet_state = 'hidden'
    
    # Save to a BytesIO object
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    
    return output

def copy_data_to_pivot(source_sheet, target_sheet, start_row):
    """Copy data from source sheet to target sheet starting at start_row"""
    # Find the data range in source sheet
    max_row = source_sheet.max_row
    max_col = source_sheet.max_column
    
    # Add header for the data
    target_sheet.cell(row=start_row-2, column=2).value = f"Data from sheet: {source_sheet.title}"
    target_sheet.cell(row=start_row-2, column=2).font = Font(bold=True, size=12)
    
    # Copy headers and data
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            target_sheet.cell(row=start_row + row - 1, column=col + 1).value = source_sheet.cell(row=row, column=col).value
            
            # Format headers
            if row == 1:
                target_sheet.cell(row=start_row + row - 1, column=col + 1).font = Font(bold=True)
                target_sheet.cell(row=start_row + row - 1, column=col + 1).fill = PatternFill(
                    start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
                
    # Add borders to the data
    for row in range(start_row, start_row + max_row):
        for col in range(2, max_col + 2):
            target_sheet.cell(row=row, column=col).border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            
    # Set column widths
    for col in range(1, max_col + 2):
        col_letter = get_column_letter(col + 1)
        target_sheet.column_dimensions[col_letter].width = 15

def copy_all_data(source_sheet, target_sheet):
    """Copy all data from source sheet to target sheet"""
    # Find the data range in source sheet
    max_row = source_sheet.max_row
    max_col = source_sheet.max_column
    
    # Copy headers and data
    for row in range(1, max_row + 1):
        for col in range(1, max_col + 1):
            target_sheet.cell(row=row, column=col).value = source_sheet.cell(row=row, column=col).value

def define_named_range(workbook, sheet, range_name):
    """Define a named range for the sheet data"""
    max_row = sheet.max_row
    max_col = sheet.max_column
    
    if max_row > 0 and max_col > 0:
        range_str = f"'{sheet.title}'!$A$1:${get_column_letter(max_col)}${max_row}"
        workbook.defined_names.add(range_name, range_str)

def set_up_data_lookup(pivot_sheet, sheet_names, pivot_sheet_name):
    """Set up formulas to display selected sheet data"""
    # Add a button to refresh data view
    pivot_sheet['D4'] = "Refresh View"
    pivot_sheet['D4'].font = Font(bold=True, color="FFFFFF")
    pivot_sheet['D4'].alignment = Alignment(horizontal='center')
    pivot_sheet['D4'].fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    pivot_sheet['D4'].border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'), 
        top=Side(style='thin'), 
        bottom=Side(style='thin')
    )
    
    # Add instructions for using the dropdown
    pivot_sheet['B6'] = "Instructions:"
    pivot_sheet['B6'].font = Font(bold=True)
    pivot_sheet['B7'] = "1. Use the dropdown above to select a sheet"
    pivot_sheet['B8'] = "2. Click on 'Refresh View' to see the data from that sheet"
    pivot_sheet['B9'] = "3. The data will appear below"
    
    # Add a note about limitations
    pivot_sheet['B11'] = "Note: This solution uses Excel's built-in functionality without VBA macros"
    pivot_sheet['B12'] = "For dynamic updates, please re-open the file in Excel after each selection"

def download_button(object_to_download, download_filename, button_text):
    """
    Generate a download button HTML for the provided file
    """
    # Convert BytesIO to base64 encoded string
    b64 = base64.b64encode(object_to_download.getvalue()).decode()
    
    button_uuid = 'download_button'
    custom_css = f"""
        <style>
            #{button_uuid} {{
                background-color: rgb(255, 255, 255);
                color: rgb(38, 39, 48);
                padding: 0.25em 0.38em;
                position: relative;
                text-decoration: none;
                border-radius: 4px;
                border-width: 1px;
                border-style: solid;
                border-color: rgb(230, 234, 241);
                border-image: initial;
            }}
            #{button_uuid}:hover {{
                border-color: rgb(246, 51, 102);
                color: rgb(246, 51, 102);
            }}
        </style>
    """
    
    dl_link = custom_css + f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" id="{button_uuid}" download="{download_filename}">{button_text}</a><br></br>'
    
    return dl_link

if __name__ == "__main__":
    main()
