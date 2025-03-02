import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
import streamlit as st
import io
import base64

def create_download_link(file, filename):
    """Generate a download link for a file"""
    b64 = base64.b64encode(file).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'
    return href

def add_data_preview(sheet, data_df, start_row):
    """Add a preview of dataframe data to the sheet"""
    # Add headers
    for col_idx, col_name in enumerate(data_df.columns, 1):
        cell = sheet.cell(row=start_row, column=col_idx, value=col_name)
        cell.font = Font(bold=True)
        cell.fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    
    # Add data (limit to first 10 rows for preview)
    preview_data = data_df.head(10)
    for row_idx, row_data in enumerate(preview_data.values, 1):
        for col_idx, value in enumerate(row_data, 1):
            sheet.cell(row=start_row + row_idx, column=col_idx, value=value)

def create_navigator_excel(uploaded_file):
    """Create a new Excel with navigation sheet and data previews"""
    
    # Read the uploaded Excel file
    file_data = io.BytesIO(uploaded_file.getvalue())
    
    # Load existing workbook
    wb = openpyxl.load_workbook(file_data)
    sheet_names = wb.sheetnames
    
    # Read sheet data for previews
    excel_data = pd.read_excel(file_data, sheet_name=None)
    
    # Create a new sheet at the beginning
    navigator_sheet = wb.create_sheet("Navigator", 0)
    
    # Set up the Navigator sheet
    navigator_sheet['A1'] = "SHEET NAVIGATOR"
    navigator_sheet['A1'].font = Font(size=16, bold=True)
    navigator_sheet.merge_cells('A1:E1')
    navigator_sheet['A1'].alignment = Alignment(horizontal='center')
    
    # Add instructions
    navigator_sheet['A3'] = "Select a sheet to view:"
    navigator_sheet['A3'].font = Font(bold=True)
    
    # Create dropdown selection cell
    sheet_list_cell = 'C3'
    navigator_sheet[sheet_list_cell] = sheet_names[1] if len(sheet_names) > 1 else ""
    
    # Create data validation (dropdown) for sheet selection
    sheet_list_str = ','.join([f'"{name}"' for name in sheet_names if name != "Navigator"])
    dv = DataValidation(type="list", formula1=f"=\"{sheet_list_str}\"", allow_blank=False)
    navigator_sheet.add_data_validation(dv)
    dv.add(sheet_list_cell)
    
    # Add a button style
    navigator_sheet['E3'] = "GO"
    button_cell = navigator_sheet['E3']
    button_cell.font = Font(color="FFFFFF", bold=True)
    button_cell.fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    button_cell.alignment = Alignment(horizontal='center')
    button_cell.border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Set column widths
    for col in range(1, 6):
        navigator_sheet.column_dimensions[get_column_letter(col)].width = 15
    
    # Add sheet buttons for quick navigation
    row = 5
    navigator_sheet.cell(row=row, column=1, value="Quick Navigation:")
    navigator_sheet.cell(row=row, column=1).font = Font(bold=True)
    
    # Add buttons for each sheet
    row += 1
    col = 1
    button_fill = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    button_font = Font(color="FFFFFF", bold=True)
    
    for sheet_name in sheet_names:
        if sheet_name != "Navigator":
            cell = navigator_sheet.cell(row=row, column=col, value=sheet_name)
            cell.fill = button_fill
            cell.font = button_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = Border(
                left=Side(style='thin'), 
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            cell.hyperlink = f"#{sheet_name}!A1"
            
            col += 1
            if col > 5:
                col = 1
                row += 1
    
    # Add section for data preview
    preview_row = row + 2
    navigator_sheet.cell(row=preview_row, column=1, value="DATA PREVIEW:")
    navigator_sheet.cell(row=preview_row, column=1).font = Font(size=12, bold=True)
    navigator_sheet.merge_cells(f'A{preview_row}:E{preview_row}')
    
    # Get first sheet data for initial preview
    if len(sheet_names) > 1:
        first_data_sheet = sheet_names[1] if sheet_names[0] == "Navigator" else sheet_names[0]
        if first_data_sheet in excel_data:
            preview_df = excel_data[first_data_sheet]
            add_data_preview(navigator_sheet, preview_df, preview_row + 2)
    
    # Hide all sheets except Navigator
    for sheet_name in sheet_names:
        if sheet_name != "Navigator":
            wb[sheet_name].sheet_state = 'hidden'
    
    # Save to a BytesIO object for download
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

# Create Streamlit app
def main():
    st.title("Excel Sheet Navigator Creator")
    
    st.write("""
    ## Upload your Excel file
    
    This tool will create a new Excel file with:
    - A Navigator sheet at the beginning
    - A dropdown menu to select sheets
    - Quick navigation buttons for each sheet
    - A data preview of the selected sheet
    - All other sheets initially hidden
    
    *No VBA or macros are used in this solution*
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Display file info
        file_details = {"Filename": uploaded_file.name, "File size": f"{uploaded_file.size/1024:.1f} KB"}
        st.write(file_details)
        
        try:
            # Check sheet count
            xls = pd.ExcelFile(uploaded_file)
            sheet_count = len(xls.sheet_names)
            st.write(f"Found {sheet_count} sheets in the uploaded file.")
            
            # Process the file
            with st.spinner("Creating Navigator Excel file..."):
                excel_file = create_navigator_excel(uploaded_file)
            
            # Create download button
            st.markdown(
                create_download_link(excel_file.getvalue(), f"Navigator_{uploaded_file.name}"),
                unsafe_allow_html=True
            )
            st.success("✅ Your Excel file with navigation has been created!")
            
            st.write("""
            ### Usage Instructions:
            1. Download the Excel file
            2. Open it in Excel
            3. On the Navigator sheet:
               - Use the dropdown to select a sheet
               - Click on sheet buttons for direct navigation
               - View data preview of selected sheets
            4. When you navigate to a sheet, it will become visible while others remain hidden
            """)
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.error("Please make sure your Excel file is not corrupted and try again.")

if __name__ == "__main__":
    main()
