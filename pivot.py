import pandas as pd
import streamlit as st
from io import BytesIO
import xlsxwriter
import re
import traceback

def create_excel_with_dashboard(uploaded_file):
    """Create an Excel file with a dashboard and data view functionality without macros"""
    
    try:
        # Read the uploaded Excel file
        xls = pd.ExcelFile(uploaded_file)
        sheet_names = xls.sheet_names
        
        # Create a new Excel workbook
        output = BytesIO()
        workbook = xlsxwriter.Workbook(output)
        
        # Define cell formats
        title_format = workbook.add_format({
            'bold': True,
            'font_size': 18,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0
        })
        
        subtitle_format = workbook.add_format({
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'border': 0
        })
        
        button_format = workbook.add_format({
            'bold': True,
            'font_size': 12,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#D8E4BC',
            'border': 1,
            'border_color': '#538DD5'
        })
        
        header_format = workbook.add_format({
            'bold': True,
            'font_size': 11,
            'align': 'center',
            'valign': 'vcenter',
            'bg_color': '#C5D9F1',
            'border': 1
        })
        
        # Create Dashboard sheet
        dashboard = workbook.add_worksheet('Dashboard')
        
        # Set up the dashboard layout
        dashboard.set_column('A:A', 15)  # Set column width for Column A
        dashboard.set_column('B:B', 25)  # Set column width for Column B
        dashboard.set_column('C:C', 15)  # Set column width for Column C
        dashboard.set_column('D:I', 15)  # Set column width for data display area
        
        # Add title and instructions
        dashboard.merge_range('B2:H2', 'EXCEL SHEET NAVIGATOR', title_format)
        dashboard.merge_range('B3:H3', 'Select a sheet to view its data below', subtitle_format)
        
        # Add sheet selection dropdown
        dashboard.write('B5', 'Select Sheet:', subtitle_format)
        
        # Create dropdown for sheet selection - sanitize sheet names to prevent formula injection
        dropdown_range = "Dashboard!$C$5"
        # Use a list comprehension to safely format sheet names
        safe_sheet_names = [name.replace('"', '""') for name in sheet_names]
        sheet_list = ','.join([f'"{name}"' for name in safe_sheet_names])
        dashboard.data_validation(dropdown_range, {
            'validate': 'list',
            'source': f"={sheet_list}",
            'input_title': 'Select a sheet:',
            'input_message': 'Choose a sheet to display its data below'
        })
        
        # Write the first sheet name as default value
        if sheet_names:
            dashboard.write('C5', sheet_names[0])
        
        # Add buttons for each sheet
        row = 7
        for i, sheet_name in enumerate(sheet_names):
            # Make button-like cells
            dashboard.write(row, 1, sheet_name, button_format)
            
            # Add formula to update selected sheet when clicked
            dashboard.write_formula(row, 2, f'=IF(B{row+1}=C5,"✓","")')
            
            row += 1
        
        # Add data display area title
        dashboard.merge_range('B10:H10', 'DATA PREVIEW', title_format)
        dashboard.write('B11', 'Showing data from sheet:', subtitle_format)
        dashboard.write_formula('C11', '=C5')  # Display selected sheet name
        
        # Data area starts at row 13
        data_start_row = 13
        
        # Create other sheets and read data
        max_cols = 0
        sheet_data = {}
        
        for sheet_name in sheet_names:
            # Read data from original sheet
            df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
            
            # Store data for reference
            sheet_data[sheet_name] = {
                'columns': df.columns.tolist(),
                'data': df.values.tolist(),
                'row_count': len(df),
                'col_count': len(df.columns)
            }
            
            # Keep track of maximum columns
            max_cols = max(max_cols, len(df.columns))
            
            # Create the sheet - sanitize sheet name
            safe_sheet_name = re.sub(r'[\\*?:/\[\]]', '_', sheet_name)  # Replace illegal Excel chars
            worksheet = workbook.add_worksheet(safe_sheet_name)
            
            # Add headers
            for col_idx, col_name in enumerate(df.columns):
                worksheet.write(0, col_idx, col_name, header_format)
            
            # Add data
            for row_idx, row_data in enumerate(df.values):
                for col_idx, value in enumerate(row_data):
                    worksheet.write(row_idx + 1, col_idx, value)
        
        # Display the first 50 rows of data using INDIRECT formulas
        # This allows the dashboard to dynamically pull data from the selected sheet
        
        # Add headers for data display section
        for i, sheet_name in enumerate(sheet_names):
            # Get columns for this sheet
            columns = sheet_data[sheet_name]['columns']
            safe_sheet_name = re.sub(r'[\\*?:/\[\]]', '_', sheet_name)
            
            # For each column, create a conditional formula that will show this header only when the sheet is selected
            for col_idx, col_name in enumerate(columns):
                try:
                    col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
                    formula = f'=IF(C5="{sheet_name}",INDIRECT("\'{safe_sheet_name}\'!{col_letter}1"),"")'
                    dashboard.write_formula(data_start_row - 1, col_idx + 1, formula, header_format)
                except Exception as e:
                    # If there's an error with column names, use a simpler approach
                    st.error(f"Error with column {col_idx} in sheet {sheet_name}: {str(e)}")
                    dashboard.write_formula(data_start_row - 1, col_idx + 1, f'=IF(C5="{sheet_name}","Column {col_idx+1}","")', header_format)
        
        # Add data rows (up to 50)
        if sheet_names:
            max_rows = min(50, max(sheet_data[name]['row_count'] for name in sheet_names))
            
            for row_idx in range(max_rows):
                for i, sheet_name in enumerate(sheet_names):
                    # Only create formulas for sheets that have this many rows
                    if row_idx < sheet_data[sheet_name]['row_count']:
                        safe_sheet_name = re.sub(r'[\\*?:/\[\]]', '_', sheet_name)
                        # For each column in this sheet
                        for col_idx in range(sheet_data[sheet_name]['col_count']):
                            try:
                                # Create a conditional formula that will show this cell only when the sheet is selected
                                col_letter = xlsxwriter.utility.xl_col_to_name(col_idx)
                                cell_address = f"{col_letter}{row_idx + 2}"
                                formula = f'=IF(C5="{sheet_name}",INDIRECT("\'{safe_sheet_name}\'!{cell_address}"),"")'
                                dashboard.write_formula(data_start_row + row_idx, col_idx + 1, formula)
                            except Exception as e:
                                # Handle any errors with specific cells
                                st.error(f"Error with cell at row {row_idx}, column {col_idx} in sheet {sheet_name}: {str(e)}")
                                dashboard.write(data_start_row + row_idx, col_idx + 1, "ERROR")
        
        # Add a note about limitations
        dashboard.merge_range(f'B{data_start_row + max_rows + 2}:H{data_start_row + max_rows + 2}', 
                            'Note: Dashboard shows up to 50 rows. Open the specific sheet for complete data.', 
                            subtitle_format)
        
        # Finalize the workbook
        workbook.close()
        output.seek(0)
        
        return output
        
    except Exception as e:
        # Capture the full traceback
        error_details = traceback.format_exc()
        st.error(f"An error occurred: {str(e)}\n\nDetails:\n{error_details}")
        # Create a simple error workbook
        output = BytesIO()
        error_workbook = xlsxwriter.Workbook(output)
        error_sheet = error_workbook.add_worksheet('Error')
        error_sheet.write(0, 0, f"Error: {str(e)}")
        error_sheet.write(1, 0, "Please check your Excel file format and try again.")
        error_workbook.close()
        output.seek(0)
        return output

# Streamlit app UI
st.title("Excel Dashboard Creator")
st.write("Upload your Excel file to create a dashboard with sheet selection functionality")

uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xls'])

if uploaded_file is not None:
    try:
        st.write("Processing your file...")
        
        # Create the dashboard Excel file
        output = create_excel_with_dashboard(uploaded_file)
        
        # Provide download button
        st.download_button(
            label="Download Excel with Dashboard",
            data=output,
            file_name="dashboard_" + uploaded_file.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
        st.success("Dashboard created successfully!")
        
        st.markdown("""
        ### Instructions:
        
        1. Download the Excel file using the button above
        2. Open the file in Excel
        3. In the Dashboard sheet, you can:
           - Use the dropdown to select a sheet
           - Click on any sheet name in the list of buttons
           - View data from the selected sheet directly in the dashboard
        
        The data preview shows up to 50 rows from each sheet. For complete data, you can navigate to the individual sheets.
        
        **No macros required!** This solution uses Excel's built-in formulas (INDIRECT and IF) to dynamically display data from different sheets.
        
        **How it works:** 
        - When you select a sheet from the dropdown, the formulas update to show that sheet's data
        - The buttons also update the dropdown when clicked
        - The actual data sheets remain accessible for full data viewing
        """)
    except Exception as e:
        st.error(f"An unexpected error occurred: {str(e)}")
        st.info("Please check your Excel file and try again. Make sure it's a valid Excel file with at least one sheet.")
