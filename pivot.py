import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.drawing.image import Image
from openpyxl.utils import get_column_letter
import streamlit as st
import io
import base64

def create_download_link(file, filename):
    """Generate a download link for a file"""
    b64 = base64.b64encode(file).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}">Download Excel file</a>'
    return href

def create_navigator_excel(uploaded_file):
    """Create a new Excel with navigation sheet and button links"""
    
    # Read the uploaded Excel file
    file_data = io.BytesIO(uploaded_file.getvalue())
    
    # Load existing workbook
    wb = openpyxl.load_workbook(file_data)
    sheet_names = wb.sheetnames
    
    # Create a new sheet at the beginning
    navigator_sheet = wb.create_sheet("Navigator", 0)
    
    # Set up the Navigator sheet
    navigator_sheet['A1'] = "SHEET NAVIGATOR"
    navigator_sheet['A1'].font = Font(size=16, bold=True)
    navigator_sheet.merge_cells('A1:D1')
    navigator_sheet['A1'].alignment = Alignment(horizontal='center')
    
    # Add instructions
    navigator_sheet['A3'] = "Click on a button below to view that sheet:"
    navigator_sheet['A3'].font = Font(italic=True)
    navigator_sheet.merge_cells('A3:D3')
    
    # Set column width
    for col in range(1, 5):
        navigator_sheet.column_dimensions[get_column_letter(col)].width = 20
    
    # Style for buttons
    button_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    button_font = Font(color="FFFFFF", bold=True)
    button_border = Border(
        left=Side(style='thin'), 
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Add navigation buttons for each sheet
    row = 5
    col = 1
    
    for i, sheet_name in enumerate(sheet_names):
        if sheet_name != "Navigator":  # Skip self-reference
            # Create a button (cell with formatting and hyperlink)
            cell = navigator_sheet.cell(row=row, column=col, value=sheet_name)
            cell.fill = button_fill
            cell.font = button_font
            cell.alignment = Alignment(horizontal='center')
            cell.border = button_border
            
            # Add hyperlink to the sheet
            cell.hyperlink = f"#{sheet_name}!A1"
            
            # Move to next position
            col += 1
            if col > 4:  # Wrap after 4 columns
                col = 1
                row += 1
    
    # Create data display area
    row += 2
    navigator_sheet.cell(row=row, column=1, value="SELECTED SHEET DATA:")
    navigator_sheet.cell(row=row, column=1).font = Font(bold=True)
    navigator_sheet.merge_cells(f'A{row}:D{row}')
    
    # Add no data placeholder
    row += 1
    navigator_sheet.cell(row=row, column=1, value="(Select a sheet to view data)")
    navigator_sheet.merge_cells(f'A{row}:D{row}')
    
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
    st.title("Excel Navigator Creator")
    
    st.write("""
    ## Upload your Excel file
    
    This app will create a new Excel file with:
    - A Navigator sheet at the beginning
    - Buttons to quickly navigate between sheets
    - All other sheets initially hidden
    """)
    
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    if uploaded_file is not None:
        # Display file info
        file_details = {"Filename": uploaded_file.name, "File size": f"{uploaded_file.size/1024:.1f} KB"}
        st.write(file_details)
        
        try:
            excel_file = create_navigator_excel(uploaded_file)
            
            # Create download button
            st.markdown(
                create_download_link(excel_file.getvalue(), f"navigator_{uploaded_file.name}"),
                unsafe_allow_html=True
            )
            st.success("✅ Your Excel file with navigation has been created!")
            
            st.write("""
            ### Instructions:
            1. Download the Excel file
            2. Open it in Excel
            3. Use the buttons on the Navigator sheet to move between sheets
            4. The selected sheet will be visible while others remain hidden
            """)
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
