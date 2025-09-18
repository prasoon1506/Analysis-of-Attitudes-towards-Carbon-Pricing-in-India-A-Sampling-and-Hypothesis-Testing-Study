import streamlit as st
import openpyxl
from datetime import datetime
import calendar
import pandas as pd
import base64
from openpyxl.styles import (Font, Alignment, Border, Side, PatternFill, NamedStyle, Color)
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.utils import get_column_letter

def calculate_projected_usage(full_month_plan, input_date, month):
    _, total_days = calendar.monthrange(datetime.now().year, datetime.strptime(month, "%b").month)
    day = int(input_date.split()[0])
    projected_usage = (day / total_days) * full_month_plan
    return projected_usage

def calculate_actual_usage_percentage(actual_usage, full_month_plan):
    try:
        return (actual_usage / full_month_plan) * 100 if full_month_plan != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def calculate_pro_rata_deviation(actual_usage_percentage, input_date, month):
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(month, "%b").month)
    day = int(input_date.split()[0])
    pro_rata_expectation = (day / total_days) * 100
    return actual_usage_percentage - pro_rata_expectation

def calculate_average_consumption(actual_usage, input_date):
    day = int(input_date.split()[0])
    try:
        return actual_usage / day if day != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def calculate_days_stock_available(current_stock, avg_consumption):
    try:
        return current_stock / avg_consumption if avg_consumption != 0 else 0
    except (TypeError, ZeroDivisionError):
        return 0

def write_formulas_to_excel(df, output_path, user_date):
    """Write data with Excel formulas to maintain formula visibility"""
    # Create a new workbook
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Bag Consumption Report'
    
    # Get date information for formulas
    input_day, input_month = user_date.split()
    day = int(input_day)
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(input_month, "%b").month)
    
    # Write headers
    headers = df.columns.tolist()
    for col, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=header)
        # Apply Aptos Narrow font to headers
        cell.font = Font(bold=True, color='FFFFFF', name='Aptos Narrow', size=12)
    
    # Write data with formulas for calculated columns
    for row_idx, (index, row) in enumerate(df.iterrows(), 2):
        # Write basic data columns (A to H)
        for col in range(1, 9):  # Columns A-H
            cell = ws.cell(row=row_idx, column=col, value=row.iloc[col-1])
            # Apply Aptos Narrow font to all cells
            cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column I: Projected Usage (Till date) = (Day/Total_Days) * Full_Month_Plan
        # Formula: =ROUND((day/total_days)*H{row},0)
        projected_formula = f"=ROUND(({day}/{total_days})*H{row_idx},0)"
        cell = ws.cell(row=row_idx, column=9, value=projected_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column J: Actual Usage % = (Actual_Usage/Full_Month_Plan)*100
        # Formula: =IF(H{row}=0,0,(F{row}/H{row}))
        usage_percent_formula = f"=IF(H{row_idx}=0,0,(F{row_idx}/H{row_idx}))"
        cell = ws.cell(row=row_idx, column=10, value=usage_percent_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column K: Pro Rata Deviation = Actual_Usage% - Pro_Rata_Expectation%
        # Formula: =J{row}-({day}/{total_days})
        pro_rata_formula = f"=J{row_idx}-({day}/{total_days})"
        cell = ws.cell(row=row_idx, column=11, value=pro_rata_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column L: Average Consumption = Actual_Usage/Day
        # Formula: =ROUND(IF({day}=0,0,F{row}/{day}),0)
        avg_consumption_formula = f"=ROUND(IF({day}=0,0,F{row_idx}/{day}),0)"
        cell = ws.cell(row=row_idx, column=12, value=avg_consumption_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column M: Days Stock Left (Based on Consumption) = Current_Stock/Average_Consumption
        # Formula: =ROUND(IF(L{row}=0,0,G{row}/L{row}),0)
        days_stock_consumption_formula = f"=ROUND(IF(L{row_idx}=0,0,G{row_idx}/L{row_idx}),0)"
        cell = ws.cell(row=row_idx, column=13, value=days_stock_consumption_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
        
        # Column N: Days Stock Left (Based on Planning) = (Opening_Balance + Full_Month_Plan - Actual_Usage)/Average_Consumption
        # Formula: =ROUND(IF(L{row}=0,0,(D{row}+H{row}-F{row})/L{row}),0)
        days_stock_planning_formula = f"=ROUND(IF(L{row_idx}=0,0,(D{row_idx}+H{row_idx}-F{row_idx})/L{row_idx}),0)"
        cell = ws.cell(row=row_idx, column=14, value=days_stock_planning_formula)
        cell.font = Font(name='Aptos Narrow', size=11)
    
    # Apply styling
    style_excel_with_formulas(wb, ws)
    
    # Save the workbook
    wb.save(output_path)

def style_excel_with_formulas(workbook, worksheet):
    """Apply styling to the Excel worksheet"""
    colors = {
        'dark_blue': '1E4C7B',
        'light_blue': '4A90E2',
        'header_bg': '2C3E50',
        'alternate_row': 'F0F8FF',
        'text_primary': '2C3E50',
        'text_secondary': '34495E'
    }
    
    border = Border(
        left=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        right=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        top=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        bottom=Side(style='thin', color=Color(rgb=colors['light_blue']))
    )
    
    header_fill = PatternFill(
        start_color=colors['header_bg'], 
        end_color=colors['header_bg'],
        fill_type='solid'
    )
    
    alternate_fill = PatternFill(
        start_color=colors['alternate_row'],
        end_color=colors['alternate_row'],
        fill_type='solid'
    )
    
    # Style headers with text wrapping
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color='FFFFFF', name='Aptos Narrow', size=12)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    
    # Style data rows
    for row in worksheet.iter_rows(min_row=2):
        for col_idx, cell in enumerate(row, 1):
            # Ensure all cells have Aptos Narrow font
            if cell.font.name != 'Aptos Narrow':
                cell.font = Font(name='Aptos Narrow', size=11)
            
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            if cell.row % 2 == 0:
                cell.fill = alternate_fill
            
            # Apply number formatting - no decimals for all numbers
            if col_idx in [4, 5, 6, 7, 8, 9, 12, 13, 14]:  # All number columns without decimals
                cell.number_format = '0'
            elif col_idx in [10, 11]:  # Percentage columns (Actual Usage % and Pro Rata Deviation)
                cell.number_format = '0%'
    
    # Auto-adjust column widths (reduced for wrapped headers)
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                # For header row, consider wrapped text
                if cell.row == 1:
                    lines = str(cell.value).split('\n') if cell.value else ['']
                    cell_length = max(len(line) for line in lines)
                else:
                    cell_length = len(str(cell.value)) if cell.value else 0
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        # Reduce width for better appearance with wrapped headers
        adjusted_width = min(max_length + 2, 25)  # Reduced max width
        worksheet.column_dimensions[column].width = adjusted_width
    
    # Set header row height for better wrapped text display
    worksheet.row_dimensions[1].height = 60
    
    # Freeze panes
    worksheet.freeze_panes = worksheet['B2']

def style_excel(df, output_path):
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    df.to_excel(writer, index=False, sheet_name='Bag Consumption Report')
    workbook = writer.book
    worksheet = writer.sheets['Bag Consumption Report']
    
    colors = {
        'dark_blue': '1E4C7B',
        'light_blue': '4A90E2',
        'header_bg': '2C3E50',
        'alternate_row': 'F0F8FF',
        'text_primary': '2C3E50',
        'text_secondary': '34495E'
    }
    
    border = Border(
        left=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        right=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        top=Side(style='thin', color=Color(rgb=colors['light_blue'])),
        bottom=Side(style='thin', color=Color(rgb=colors['light_blue']))
    )
    
    header_fill = PatternFill(
        start_color=colors['header_bg'], 
        end_color=colors['header_bg'],
        fill_type='solid'
    )
    
    alternate_fill = PatternFill(
        start_color=colors['alternate_row'],
        end_color=colors['alternate_row'],
        fill_type='solid'
    )
    
    # Style headers with text wrapping
    for cell in worksheet[1]:
        cell.font = Font(bold=True, color='FFFFFF', name='Aptos Narrow', size=12)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border
    
    # Style data rows
    for row in worksheet.iter_rows(min_row=2):
        for col_idx, cell in enumerate(row, 1):
            cell.font = Font(name='Aptos Narrow', size=11)
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = border
            if cell.row % 2 == 0:
                cell.fill = alternate_fill
            
            # Apply number formatting
            if col_idx in [4, 5, 6, 7, 9, 12, 13, 14]:  # Number columns without decimals
                cell.number_format = '0'
            elif col_idx in [10, 11]:  # Percentage columns (Actual Usage % and Pro Rata Deviation)
                cell.number_format = '0%'
    
    # Auto-adjust column widths (reduced for wrapped headers)
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter
        for cell in col:
            try:
                # For header row, consider wrapped text
                if cell.row == 1:
                    lines = str(cell.value).split('\n') if cell.value else ['']
                    cell_length = max(len(line) for line in lines)
                else:
                    cell_length = len(str(cell.value)) if cell.value else 0
                if cell_length > max_length:
                    max_length = cell_length
            except:
                pass
        # Reduce width for better appearance with wrapped headers
        adjusted_width = min(max_length + 2, 25)  # Reduced max width
        worksheet.column_dimensions[column].width = adjusted_width
    
    # Set header row height for better wrapped text display
    worksheet.row_dimensions[1].height = 60
    
    worksheet.freeze_panes = worksheet['B2']
    worksheet.title = 'Bag Consumption Report'
    writer.close()

def filter_and_rename_columns(input_file, merge_file, user_date):
    wb = openpyxl.load_workbook(input_file)
    ws = wb.active
    
    merge_wb = openpyxl.load_workbook(merge_file)
    merge_ws = merge_wb.active
    
    merge_data = {}
    for row in merge_ws.iter_rows(min_row=2):
        key = (
            str(row[0].value).strip() if row[0].value is not None else '', 
            str(row[1].value).strip() if row[1].value is not None else '',
            str(row[2].value).strip() if row[2].value is not None else ''
        )
        merge_data[key] = str(row[3].value).strip() if row[3].value is not None else ''
    
    columns_to_keep = [1, 2, 3, 4, 6, 8, 9]
    data_rows = []
    header = []
    
    input_day, input_month = user_date.split()
    _, total_days = calendar.monthrange(datetime.now().year, 
        datetime.strptime(input_month, "%b").month)
    
    for row_num, row in enumerate(ws.iter_rows(min_row=1, max_row=ws.max_row), 1):
        new_row = []
        row_values = []
        
        for idx, cell in enumerate(row, 1):
            if idx in columns_to_keep:
                new_row.append(cell.value)
                if idx <= 3:
                    row_values.append(str(cell.value).strip() if cell.value is not None else '')
        
        row_match_key = tuple(row_values)
        
        if row_num == 1:
            header = [
                "Plant\nName",
                "Brand\nName", 
                "Bag\nName", 
                "Opening Balance\nas on 01.09.2025",
                "Tomonth\nReceipt",
                f"Actual Usage\n(Till {user_date})",
                "Current\navailable stock",
                "Full Month\nPlan",
                f"Projected Usage\n(Till {user_date})",
                f"Actual Usage %\n(Till {user_date})\n(Based on Planning)",
                "Pro Rata\nDeviation",
                "Average\nConsumption",
                "No. of Days\nStock Left\n(Based on Consumption)",
                "No. of Days\nStock Left\n(Based on Planning)"
            ]
            continue
        
        full_month_plan = merge_data.get(row_match_key, "0")
        try:
            full_month_plan = float(full_month_plan)
        except ValueError:
            full_month_plan = 0
        
        new_row.append(full_month_plan)
        
        projected_usage = calculate_projected_usage(full_month_plan, user_date, input_month)
        new_row.append(int(projected_usage))
        
        actual_usage = new_row[5]
        opening_balance = new_row[3]
        current_stock = new_row[6]
        
        actual_usage_percentage = calculate_actual_usage_percentage(actual_usage, full_month_plan)
        new_row.append(int(actual_usage_percentage))
        
        pro_rata_deviation = calculate_pro_rata_deviation(actual_usage_percentage, user_date, input_month)
        new_row.append(int(pro_rata_deviation))
        
        average_consumption = calculate_average_consumption(actual_usage, user_date)
        new_row.append(int(average_consumption) if average_consumption is not None else 0)
        
        days_stock_tomonth_receipt = calculate_days_stock_available(current_stock, average_consumption)
        new_row.append(int(days_stock_tomonth_receipt))
        
        try:
            days_stock_planning = calculate_days_stock_available(
                opening_balance + full_month_plan - actual_usage, 
                average_consumption
            )
        except (TypeError, ValueError):
            days_stock_planning = 0
        new_row.append(int(days_stock_planning))
        
        if isinstance(new_row[0], str) and len(new_row[0]) > 9:
            new_row[0] = new_row[0][9:]
        if isinstance(new_row[2], str) and len(new_row[2]) > 9:
            new_row[2] = new_row[2][9:]
        
        # Skip rows where Plant Name is "Totals" (case insensitive)
        if isinstance(new_row[0], str) and new_row[0].lower().strip() == "totals":
            continue
        
        data_rows.append(new_row)
    
    df = pd.DataFrame(data_rows, columns=header)
    return df

def get_download_link(df, user_date, use_formulas=True):
    output_path = 'bag_report.xlsx'
    
    if use_formulas:
        write_formulas_to_excel(df, output_path, user_date)
    else:
        style_excel(df, output_path)
    
    with open(output_path, 'rb') as f:
        bytes = f.read()
    b64 = base64.b64encode(bytes).decode()
    href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="bag_report.xlsx">Download Professional Excel Report with Formulas</a>'
    return href

def main():
    st.set_page_config(page_title="Bag Report", layout="wide", page_icon="📊")
    
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
    
    st.markdown('<h1 style="color:#1E4C7B; text-align:center; border-bottom: 3px solid #4A90E2; padding-bottom: 10px;">📊 Bag Consumption Report</h1>', unsafe_allow_html=True)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        input_file = st.file_uploader("Upload Input Excel File", type=['xlsx', 'xls'], help="Select the input file for analysis")
    
    with col2:
        merge_file = st.file_uploader("Upload Merge Excel File", type=['xlsx', 'xls'], help="Select the merge file for additional data")
    
    with col3:
        user_date = st.text_input("Enter Date (e.g., 01 Sep)", value="01 Sep", help="Enter the date for report generation")
    
    # Add option for formula inclusion
    st.markdown("---")
    use_formulas = st.checkbox("Include Excel Formulas in Report", value=True, help="When checked, formulas will be visible when clicking on calculated cells in Excel")
    
    if use_formulas:
        st.info("✅ Formulas will be included - you can see and copy formulas by clicking on calculated cells in Excel")
    else:
        st.warning("⚠️ Only calculated values will be included - formulas won't be visible")
    
    if input_file and merge_file and user_date:
        try:
            with open("input_file.xlsx", "wb") as f:
                f.write(input_file.getbuffer())
            with open("merge_file.xlsx", "wb") as f:
                f.write(merge_file.getbuffer())
            
            df = filter_and_rename_columns("input_file.xlsx", "merge_file.xlsx", user_date)
            
            start_date = "01 Sep"
            st.markdown(f'<p class="big-font">Consumption Analysis</p>', unsafe_allow_html=True)
            st.markdown(f'<p class="sub-font">Period: {start_date} to {user_date}</p>', unsafe_allow_html=True)
            
            st.dataframe(df, use_container_width=True)
            
            st.markdown(get_download_link(df, user_date, use_formulas), unsafe_allow_html=True)
            
            if use_formulas:
                st.markdown("### Formula Details")
                st.markdown("""
                **The following formulas are used in the calculated columns:**
                
                - **Projected Usage**: `=ROUND((Day/Total_Days) * Full_Month_Plan, 0)` (No decimals)
                - **Actual Usage %**: `=IF(Full_Month_Plan=0, 0, Actual_Usage/Full_Month_Plan)` (Percentage format - displays as %)
                - **Pro Rata Deviation**: `=Actual_Usage% - (Day/Total_Days)` (Percentage format - displays as %)
                - **Average Consumption**: `=ROUND(IF(Day=0, 0, Actual_Usage/Day), 0)` (No decimals)
                - **Days Stock Left (Consumption)**: `=ROUND(IF(Avg_Consumption=0, 0, Current_Stock/Avg_Consumption), 0)` (No decimals)
                - **Days Stock Left (Planning)**: `=ROUND(IF(Avg_Consumption=0, 0, (Opening_Balance+Full_Month_Plan-Actual_Usage)/Avg_Consumption), 0)` (No decimals)
                
                **Formatting Applied:**
                - **Font**: Entire report uses Aptos Narrow font
                - **Numbers**: All numeric values display without decimal points
                - **Percentages**: Columns 10 & 11 show values in percentage format (e.g., 45% instead of 0.45)
                """)
            
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")

if __name__ == "__main__":
    main()
