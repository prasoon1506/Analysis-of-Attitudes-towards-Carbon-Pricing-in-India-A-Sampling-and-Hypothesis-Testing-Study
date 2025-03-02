import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.dimensions import ColumnDimension, DimensionHolder
from openpyxl.utils import get_column_letter
import io
import streamlit as st
from io import BytesIO

def create_dashboard_excel(uploaded_file):
    # Read the Excel file with pandas to get sheet names
    xls = pd.ExcelFile(uploaded_file)
    sheet_names = xls.sheet_names
    
    # Now use openpyxl for more advanced manipulation
    wb = openpyxl.load_workbook(uploaded_file)
    
    # Create a new Dashboard sheet and insert it at the beginning
    if 'Dashboard' in wb.sheetnames:
        # If Dashboard already exists, remove it first
        dashboard_sheet_index = wb.sheetnames.index('Dashboard')
        wb.remove(wb[wb.sheetnames[dashboard_sheet_index]])
    
    # Create and insert the Dashboard sheet at the beginning
    wb.create_sheet('Dashboard', 0)
    dashboard = wb['Dashboard']
    
    # Set up styles for the dashboard
    title_font = Font(name='Arial', size=16, bold=True)
    button_font = Font(name='Arial', size=12, bold=True)
    button_fill = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    
    # Add title
    dashboard['A1'] = "Excel Sheet Navigator"
    dashboard['A1'].font = title_font
    dashboard.merge_cells('A1:E1')
    dashboard['A1'].alignment = Alignment(horizontal='center')
    
    # Add instructions
    dashboard['A3'] = "Click on a button below to navigate to the respective sheet:"
    dashboard['A3'].font = Font(name='Arial', size=12)
    dashboard.merge_cells('A3:E3')
    
    # Create buttons for each sheet
    row = 5
    for i, sheet_name in enumerate(sheet_names):
        if sheet_name != 'Dashboard':  # Skip the dashboard itself
            button_cell = f'B{row}'
            dashboard[button_cell] = sheet_name
            dashboard[button_cell].font = button_font
            dashboard[button_cell].fill = button_fill
            dashboard[button_cell].alignment = Alignment(horizontal='center')
            
            # Add border to make it look like a button
            thin_border = Border(left=Side(style='thin'), 
                                right=Side(style='thin'), 
                                top=Side(style='thin'), 
                                bottom=Side(style='thin'))
            dashboard[button_cell].border = thin_border
            
            row += 2
    
    # Adjust column widths
    dim_holder = DimensionHolder(worksheet=dashboard)
    for col in range(1, 10):
        dim_holder[get_column_letter(col)] = ColumnDimension(dashboard, min=col, width=20)
    dashboard.column_dimensions = dim_holder
    
    # Add VBA code for sheet navigation
    # Create a macro-enabled workbook
    if not wb.vba_archive:
        wb.create_vba_module()
    
    # Add a VBA module for the navigation buttons
    vba_code = """
    Sub NavigateToSheet(sheetName As String)
        ' Hide all sheets except Dashboard and target sheet
        Dim ws As Worksheet
        For Each ws In ActiveWorkbook.Worksheets
            If ws.Name <> "Dashboard" And ws.Name <> sheetName Then
                ws.Visible = xlSheetHidden
            Else
                ws.Visible = xlSheetVisible
            End If
        Next ws
        
        ' Activate the selected sheet
        Worksheets(sheetName).Activate
    End Sub
    """
    
    # Add shape buttons with macro assignments
    for i, sheet_name in enumerate(sheet_names):
        if sheet_name != 'Dashboard':
            button_row = 5 + i*2
            dashboard_sheet = wb['Dashboard']
            
            # Add a hyperlink that calls the VBA function
            button_cell = f'B{button_row}'
            dashboard_sheet[button_cell].hyperlink = f"#'{sheet_name}'!A1"
            
            # We'll add a comment with instructions since we can't directly add the VBA here
            dashboard_sheet[button_cell].comment = openpyxl.comments.Comment(
                "In Excel, right-click this cell, select 'Assign Macro', and create a macro that hides other sheets and shows this one", "Python Script")
    
    # Hide all sheets except Dashboard initially
    for sheet_name in wb.sheetnames:
        if sheet_name != 'Dashboard':
            wb[sheet_name].sheet_state = 'hidden'
    
    # Convert to bytes for downloading
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    
    return output

# Streamlit UI
st.title("Excel Dashboard Creator")
st.write("Upload an Excel file to add a navigation dashboard.")

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xlsm'])

if uploaded_file is not None:
    st.write("Processing your file...")
    
    # Create the dashboard Excel file
    output = create_dashboard_excel(uploaded_file)
    
    # Provide download button
    st.download_button(
        label="Download Excel with Dashboard",
        data=output,
        file_name="dashboard_" + uploaded_file.name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    
    st.success("Dashboard created successfully! Note: After downloading, you'll need to enable macros and assign the navigation macros to the buttons.")
    
    # Instructions for the user
    st.markdown("""
    ### Instructions for using the dashboard:
    
    1. Download the Excel file using the button above.
    2. Open the file in Microsoft Excel.
    3. If prompted, enable macros (this requires Excel Desktop, not Excel Online).
    4. In the Dashboard sheet, you'll see buttons for each sheet in your workbook.
    5. For each button:
       - Right-click the cell containing the sheet name
       - Select "Assign Macro"
       - Create a new macro with code like:
       ```
       Sub SheetName_Click()
           ' Hide all sheets except Dashboard and target sheet
           Dim ws As Worksheet
           For Each ws In ActiveWorkbook.Worksheets
               If ws.Name <> "Dashboard" And ws.Name <> "SheetName" Then
                   ws.Visible = xlSheetHidden
               Else
                   ws.Visible = xlSheetVisible
               End If
           Next ws
           
           ' Activate the selected sheet
           Worksheets("SheetName").Activate
       End Sub
       ```
       (Replace "SheetName" with the actual sheet name)
    
    6. Save the file as an Excel Macro-Enabled Workbook (.xlsm)
    
    Note: If you prefer not to use macros, the cells already contain hyperlinks to jump to each sheet, but they won't automatically hide other sheets.
    """)

# Alternative approach using pure Python without requiring macro assignment
st.markdown("""
### Alternative approach without macros:

If you don't want to use macros, I can also create a version that:
1. Creates a completely new Excel file with the Dashboard sheet
2. Adds a sheet for each original sheet that contains a copy of the data
3. Includes Excel formulas and shapes that allow navigation without macros

Just let me know if you'd prefer this approach!
""")
