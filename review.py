from openpyxl import load_workbook 
from openpyxl.styles import PatternFill 
import re 
def is_percentage(value): 
    if isinstance(value, str): 
        cleaned = value.strip().rstrip('%') 
        try: 
            float(cleaned) 
            return True 

        except ValueError: 

            return False 

    return isinstance(value, (int, float)) 

 

def get_numeric_value(value): 

    """Convert percentage string to numeric value.""" 

    if isinstance(value, str): 

        # Remove any whitespace and % symbol 

        return float(value.strip().rstrip('%')) 

    return float(value) 

 

def format_excel_file(input_file, output_file): 

    # Load the workbook 

    wb = load_workbook(input_file) 

    ws = wb.active 

     

    # Define color fills 

    red_fill = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid') 

    green_fill = PatternFill(start_color='FF00FF00', end_color='FF00FF00', fill_type='solid') 

     

    # Define column pairs to compare (1-based index) 

    column_pairs = [ 

        ('G', 'H'), 

        ('I', 'J'), 

        ('K', 'L'), 

        ('M', 'N') 

    ] 

     

    # Process each row 

    for row in range(2, ws.max_row + 1):  # Start from row 2 to skip header 

        for col1, col2 in column_pairs: 

            # Get cell values 

            cell1 = ws[f'{col1}{row}'] 

            cell2 = ws[f'{col2}{row}'] 

             

            # Skip if either cell is empty 

            if cell1.value is None or cell2.value is None: 

                continue 

             

            # Check if both cells contain percentage values 

            if is_percentage(cell1.value) and is_percentage(cell2.value): 

                val1 = get_numeric_value(cell1.value) 

                val2 = get_numeric_value(cell2.value) 

                 

                # Apply formatting based on comparison 

                if val2 < val1: 

                    ws[f'{col2}{row}'].fill = red_fill 

                else:  # val2 >= val1 

                    ws[f'{col2}{row}'].fill = green_fill 

     

    # Save the workbook 

    wb.save(output_file) 

    print(f"File has been processed and saved as {output_file}") 

 

if __name__ == "__main__": 

    # Example usage 

    input_file = "/content/Copy of President_Report_YTD.xlsx" 

    output_file = "output.xlsx" 

    format_excel_file(input_file, output_file) 

 
