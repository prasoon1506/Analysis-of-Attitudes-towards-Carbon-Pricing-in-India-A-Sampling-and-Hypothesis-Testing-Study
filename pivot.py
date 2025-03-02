import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO
import os
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
    pivot_sheet['A2'] = "Click on a button below to view the data from the corresponding sheet"
    
    title_font = Font(bold=True, size=14)
    instruction_font = Font(italic=True, size=12)
    pivot_sheet['A1'].font = title_font
    pivot_sheet['A2'].font = instruction_font
    
    # Add buttons for each sheet
    row = 4
    col = 2
    max_cols = 4  # Maximum columns for buttons
    
    # Create a VBA module with macros to show data
    vba_code = """
Sub ShowSheetData(SheetName As String)
    ' Clear previous data display
    ClearDataDisplay
    
    ' Copy data from selected sheet
    Dim sourceSheet As Worksheet
    Set sourceSheet = ThisWorkbook.Sheets(SheetName)
    
    ' Determine data range
    Dim lastRow As Long, lastCol As Long
    lastRow = sourceSheet.Cells(sourceSheet.Rows.Count, 1).End(xlUp).Row
    lastCol = sourceSheet.Cells(1, sourceSheet.Columns.Count).End(xlToLeft).Column
    
    ' Copy data to Navigation sheet
    Dim destRow As Long
    destRow = 12  ' Starting row for data display
    
    ' Add header with sheet name
    ActiveSheet.Cells(destRow - 2, 2) = "Data from sheet: " & SheetName
    ActiveSheet.Cells(destRow - 2, 2).Font.Bold = True
    ActiveSheet.Cells(destRow - 2, 2).Font.Size = 12
    
    ' Copy headers and data
    sourceSheet.Range(sourceSheet.Cells(1, 1), sourceSheet.Cells(lastRow, lastCol)).Copy
    ActiveSheet.Cells(destRow, 2).PasteSpecial xlPasteAll
    
    ' Format the copied data
    With ActiveSheet.Range(ActiveSheet.Cells(destRow, 2), ActiveSheet.Cells(destRow + lastRow - 1, lastCol + 1))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        With .Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217)
        End With
    End With
    
    ' Clean up
    Application.CutCopyMode = False
End Sub

Sub ClearDataDisplay()
    ' Clear data display area (row 12 and below)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Navigation")
    
    ' Check if there's any data to clear
    If ws.Cells(12, 2).Value <> "" Then
        Dim lastRow As Long, lastCol As Long
        lastRow = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If lastRow >= 12 Then
            lastCol = ws.Cells(12, ws.Columns.Count).End(xlToLeft).Column
            ws.Range(ws.Cells(10, 2), ws.Cells(lastRow, lastCol)).Clear
        End If
    End If
End Sub
    """
    
    # Create buttons for each sheet with VBA macros
    for i, sheet_name in enumerate(sheet_names):
        if sheet_name != pivot_sheet_name:
            cell = pivot_sheet.cell(row=row, column=col)
            cell.value = sheet_name
            
            # Button style
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            
            # Apply border
            thin_border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'), 
                top=Side(style='thin'), 
                bottom=Side(style='thin')
            )
            cell.border = thin_border
            
            # Add a shape button with VBA macro
            # Since we can't directly add buttons with openpyxl, we'll set up a hyperlink
            # and provide instructions on creating buttons in Excel
            cell_address = f"{get_column_letter(col)}{row}"
            
            # Adjust column width
            col_letter = get_column_letter(col)
            pivot_sheet.column_dimensions[col_letter].width = max(15, len(sheet_name) + 4)
            pivot_sheet.row_dimensions[row].height = 30
            
            # Move to next position
            col += 1
            if col > max_cols + 1:
                col = 2
                row += 2
    
    # Add a section for sheet preview
    preview_row = row + 3
    pivot_sheet.cell(row=preview_row, column=2).value = "Instructions:"
    pivot_sheet.cell(row=preview_row, column=2).font = Font(bold=True, size=12)
    
    preview_row += 1
    instructions = [
        "1. After downloading the file, open it in Excel.",
        "2. If prompted about macros, click 'Enable Macros'.",
        "3. Right-click on each sheet name above and select 'Assign Macro'.",
        f"4. Choose 'ShowSheetData' macro and add the sheet name in quotes as parameter (e.g., \"Sheet1\").",
        "5. Click on any button to display that sheet's data below."
    ]
    
    for instruction in instructions:
        pivot_sheet.cell(row=preview_row, column=2).value = instruction
        preview_row += 1
    
    # Hide all sheets except the pivot sheet
    for sheet_name in sheet_names:
        if sheet_name != pivot_sheet_name:
            workbook[sheet_name].sheet_state = 'hidden'
    
    # Create a VBA module for the workbook
    # Note: openpyxl doesn't support VBA directly, so we'll add instructions
    pivot_sheet.cell(row=preview_row + 2, column=2).value = "VBA Code to Add:"
    pivot_sheet.cell(row=preview_row + 2, column=2).font = Font(bold=True, size=12)
    
    vba_rows = vba_code.split('\n')
    for i, vba_row in enumerate(vba_rows):
        pivot_sheet.cell(row=preview_row + 3 + i, column=2).value = vba_row
    
    # Save to a BytesIO object
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    
    return output

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
