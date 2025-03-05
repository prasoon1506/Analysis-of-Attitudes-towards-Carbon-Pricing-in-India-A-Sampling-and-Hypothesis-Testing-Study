import streamlit as st
import openpyxl
from datetime import datetime
import calendar
import pandas as pd
import base64
from openpyxl.styles import (
    Font, Alignment, Border, Side, 
    PatternFill, NamedStyle
)

def calculate_projected_usage(full_month_plan, input_date, month):
    """
    Calculate projected usage based on the input date and month
    """
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(month, "%b").month)
    
    day = int(input_date.split()[0])
    
    projected_usage = (day / total_days) * full_month_plan
    return projected_usage

def calculate_actual_usage_percentage(actual_usage, full_month_plan):
    """
    Calculate actual usage percentage
    """
    try:
        return (actual_usage / full_month_plan) * 100 if full_month_plan != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def calculate_pro_rata_deviation(actual_usage_percentage, input_date, month):
    """
    Calculate pro-rata deviation
    """
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(month, "%b").month)
    
    day = int(input_date.split()[0])
    
    pro_rata_expectation = (day / total_days) * 100
    
    return actual_usage_percentage - pro_rata_expectation

def calculate_average_consumption(actual_usage, input_date):
    """
    Calculate average daily consumption
    """
    day = int(input_date.split()[0])
    
    try:
        return actual_usage / day if day != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def calculate_days_stock_available(current_stock, avg_consumption):
    """
    Calculate days of stock available based on average consumption
    """
    try:
        return current_stock / avg_consumption if avg_consumption != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def style_excel(df, output_path):
    """
    Apply professional styling to the Excel file
    """
    # Create a Pandas Excel writer using openpyxl as the engine
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    
    # Write the dataframe to the Excel file
    df.to_excel(writer, index=False, sheet_name='Bag Report')
    
    # Get the workbook and worksheet
    workbook = writer.book
    worksheet = writer.sheets['Bag Report']
    
    # Define color palette
    header_fill = PatternFill(start_color='1E4C7B', end_color='1E4C7B', fill_type='solid')
    alternate_fill = PatternFill(start_color='F0F2F6', end_color='F0F2F6', fill_type='solid')
    
    # Define border style
    border = Border(
        left=Side(style='thin', color='4A90E2'),
        right=Side(style='thin', color='4A90E2'),
        top=Side(style='thin', color='4A90E2'),
        bottom=Side(style='thin', color='4A90E2')
    )
    
    # Style header
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color='FFFFFF', name='Arial', size=12)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border
    
    # Adjust column widths and style data cells
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
            
            # Style data cells
            if cell.row > 1:  # Skip header row
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
                
                # Alternate row background
                if cell.row % 2 == 0:
                    cell.fill = alternate_fill
    
    # Adjust column widths
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width
    
    # Freeze top row
    worksheet.freeze_panes = worksheet['A2']
    
    # Save the file
    writer.close()

def filter_and_rename_columns(input_file, merge_file, user_date):
    """
    Process Excel files and return a DataFrame with calculated columns
    """
    # Load the workbooks and active sheets
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active

    # Load the merge file
    merge_wb = openpyxl.load_workbook(merge_file)
    merge_ws = merge_wb.active

    # Create a dictionary to store merge file data
    merge_data = {}
    
    # Populate merge data dictionary
    for row in merge_ws.iter_rows(min_row=2):  # Assuming first row is header
        key = (str(row[0].value).strip() if row[0].value is not None else '', 
               str(row[1].value).strip() if row[1].value is not None else '', 
               str(row[2].value).strip() if row[2].value is not None else '')
        
        merge_data[key] = str(row[3].value).strip() if row[3].value is not None else ''

    # Define the columns to keep: 1st to 4th, 6th, 8th, and 9th columns
    columns_to_keep = [1, 2, 3, 4, 6, 8, 9]

    # Prepare lists to create DataFrame
    data_rows = []
    header = []

    # Parse the user date
    input_day, input_month = user_date.split()
    
    # Parse the total days in the month
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(input_month, "%b").month)

    # Iterate over the rows in the original worksheet
    for row_num, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
        new_row = []
        row_values = []
        
        # Collect values for the first 3 columns
        for idx, cell in enumerate(row, 1):
            if idx in columns_to_keep:
                new_row.append(cell.value)
                
                # For first 3 columns, store for potential matching
                if idx <= 3:
                    # Convert to string and strip whitespace, handle None
                    row_values.append(str(cell.value).strip() if cell.value is not None else '')
        
        # Create a key for matching
        row_match_key = tuple(row_values)
        
        # If first row (header), save the header
        if row_num == 1:
            header = [
                "Material Description", 
                "Code", 
                "Issue", 
                "Opening Balance as on 01.03.2025", 
                "Tomonth Receipt", 
                f"Actual Usage (Till {user_date})", 
                "Current available stock",
                "Full Month Plan",
                f"Projected Usage (Till {user_date})",
                f"Actual Usage % (Till {user_date}) (Based on Planning)",
                "Pro Rata Deviation",
                "Average Consumption",
                "Days Stock (TomonthReceipt)",
                "Days Stock (Planning)"
            ]
            continue
        
        # Process data rows
        full_month_plan = merge_data.get(row_match_key, "0")
        try:
            full_month_plan = float(full_month_plan)
        except ValueError:
            full_month_plan = 0
        
        # Add Full Month Plan
        new_row.append(full_month_plan)
        
        # Calculate Projected Usage
        projected_usage = calculate_projected_usage(full_month_plan, user_date, input_month)
        new_row.append(int(projected_usage))
        
        # Get Actual Usage from the original data (6th column after filtering)
        actual_usage = new_row[5]  # This is the 'Actual Usage' column
        
        # Get Opening Balance from the original data (4th column after filtering)
        opening_balance = new_row[3]
        
        # Get Current Available Stock (8th column after filtering)
        current_stock = new_row[6]
        
        # Calculate Actual Usage Percentage
        actual_usage_percentage = calculate_actual_usage_percentage(actual_usage, full_month_plan)
        new_row.append(int(actual_usage_percentage))
        
        # Calculate Pro Rata Deviation
        pro_rata_deviation = calculate_pro_rata_deviation(actual_usage_percentage, user_date, input_month)
        new_row.append(int(pro_rata_deviation))
        
        # Calculate Average Consumption
        average_consumption = calculate_average_consumption(actual_usage, user_date)
        new_row.append(int(average_consumption) if average_consumption is not None else 0)
        
        # Calculate Days Stock Available (Based on TomonthReceipt)
        days_stock_tomonth_receipt = calculate_days_stock_available(current_stock, average_consumption)
        new_row.append(int(days_stock_tomonth_receipt))
        
        # Calculate Days Stock Available (Based on Planning)
        try:
            days_stock_planning = calculate_days_stock_available(
                opening_balance + full_month_plan - actual_usage, 
                average_consumption
            )
        except (TypeError, ValueError):
            days_stock_planning = 0
        new_row.append(int(days_stock_planning))
        
        # Remove the first 9 letters (including space) from the first two columns
        if isinstance(new_row[0], str) and len(new_row[0]) > 9:
            new_row[0] = new_row[0][9:]
        
        # Remove the first 9 letters (including space) from the third column (Issue)
        if isinstance(new_row[2], str) and len(new_row[2]) > 9:
            new_row[2] = new_row[2][9:]
        
        data_rows.append(new_row)

    # Create DataFrame
    df = pd.DataFrame(data_rows, columns=header)
    return df

def get_download_link(df):
    """
    Create a download link for the DataFrame
    """
    output_path = 'bag_report.xlsx'
    
    # Style the Excel file
    style_excel(df, output_path)
    
    # Read the file and create download link
    with open(output_path, 'rb') as f:
        bytes = f.read()
    b64 = base64.b64encode(bytes).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="bag_report.xlsx">Download Professional Excel Report</a>'
    return href

def main():
    # Set page configuration
    st.set_page_config(page_title="Bag Report", layout="wide", page_icon="📊")
    
    # Custom CSS for professional styling
    st.markdown("""
    <style>
    .reportview-container {
        background-color: #f0f2f6;
    }
    .big-font {
        font-size:24px !important;
        color: #1E4C7B;
        font-weight: bold;
        text-transform: uppercase;
        letter-spacing: 1px;
    }
    .sub-font {
        font-size:18px !important;
        color: #4A90E2;
        font-style: italic;
    }
    </style>
    """, unsafe_allow_html=True)

    # Title and Introduction
    st.markdown('<h1 style="color:#1E4C7B; text-align:center; border-bottom: 3px solid #4A90E2; padding-bottom: 10px;">📊 Bag Consumption Report</h1>', unsafe_allow_html=True)

    # File Upload Section
    col1, col2, col3 = st.columns(3)
    
    with col1:
        input_file = st.file_uploader("Upload Input Excel File", type=['xlsx', 'xls'], help="Select the input file for analysis")
    
    with col2:
        merge_file = st.file_uploader("Upload Merge Excel File", type=['xlsx', 'xls'], help="Select the merge file for additional data")
    
    with col3:
        user_date = st.text_input("Enter Date (e.g., 04 Mar)", value="04 Mar", help="Enter the date for report generation")

    # Process and Display Report
    if input_file and merge_file and user_date:
        try:
            # Save uploaded files temporarily
            with open("input_file.xlsx", "wb") as f:
                f.write(input_file.getbuffer())
            
            with open("merge_file.xlsx", "wb") as f:
                f.write(merge_file.getbuffer())
            
            # Process the files
            df = filter_and_rename_columns("input_file.xlsx", "merge_file.xlsx", user_date)
            
            # Create period and detailed title
            start_date = "01 Mar"
            st.markdown(f'<p class="big-font">Consumption Analysis</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="sub-font">Period: {start_date} to {user_date}</p>', unsafe_allow_html=True)
            
            # Display DataFrame
            st.dataframe(df, use_container_width=True)
            
            # Download Link
            st.markdown(get_download_link(df), unsafe_allow_html=True)
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
