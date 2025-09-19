import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import re
from typing import Dict, List, Optional, Union

# Configure Streamlit page
st.set_page_config(
    page_title="Cement Ledger to Excel Converter",
    page_icon="🏗️",
    layout="wide"
)

class LedgerProcessor:
    """Process different formats of cement ledger files"""
    
    def __init__(self):
        self.standard_columns = [
            'Date', 'Brand', 'Cement_Type', 'Quantity_Bags', 
            'Rate_Per_Bag', 'Total_Amount', 'Supplier', 
            'Invoice_Number', 'Vehicle_Number', 'Remarks'
        ]
    
    def detect_file_type(self, file) -> str:
        """Detect the type of uploaded file"""
        if file.name.endswith('.csv'):
            return 'csv'
        elif file.name.endswith(('.xlsx', '.xls')):
            return 'excel'
        elif file.name.endswith('.txt'):
            return 'text'
        else:
            return 'unknown'
    
    def read_file(self, file) -> pd.DataFrame:
        """Read file based on its type"""
        file_type = self.detect_file_type(file)
        
        try:
            if file_type == 'csv':
                # Try different encodings and separators
                encodings = ['utf-8', 'latin-1', 'cp1252']
                separators = [',', ';', '\t', '|']
                
                for encoding in encodings:
                    for sep in separators:
                        try:
                            file.seek(0)
                            df = pd.read_csv(file, encoding=encoding, sep=sep)
                            if len(df.columns) > 1:  # Valid CSV found
                                return df
                        except:
                            continue
                
                # If all fails, try basic read
                file.seek(0)
                return pd.read_csv(file)
                
            elif file_type == 'excel':
                return pd.read_excel(file, sheet_name=0)
                
            elif file_type == 'text':
                # For text files, try to parse as delimited
                content = file.read().decode('utf-8')
                lines = content.strip().split('\n')
                
                # Try to detect delimiter
                delimiters = ['\t', '|', ',', ';', ' ']
                for delimiter in delimiters:
                    if delimiter in lines[0]:
                        data = [line.split(delimiter) for line in lines]
                        return pd.DataFrame(data[1:], columns=data[0])
                
                # If no delimiter found, create single column
                return pd.DataFrame({'Raw_Data': lines})
            
            else:
                st.error(f"Unsupported file type: {file.name}")
                return pd.DataFrame()
                
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            return pd.DataFrame()
    
    def clean_column_names(self, df: pd.DataFrame) -> pd.DataFrame:
        """Clean and standardize column names"""
        df_clean = df.copy()
        
        # Remove extra whitespace and special characters
        df_clean.columns = df_clean.columns.str.strip().str.replace(r'[^\w\s]', '', regex=True)
        
        # Common column mappings
        column_mappings = {
            # Date columns
            'date': 'Date', 'dt': 'Date', 'transaction_date': 'Date',
            'entry_date': 'Date', 'purchase_date': 'Date',
            
            # Brand columns
            'brand': 'Brand', 'cement_brand': 'Brand', 'manufacturer': 'Brand',
            'company': 'Brand', 'brand_name': 'Brand',
            
            # Type columns
            'type': 'Cement_Type', 'cement_type': 'Cement_Type', 'grade': 'Cement_Type',
            'category': 'Cement_Type', 'specification': 'Cement_Type',
            
            # Quantity columns
            'quantity': 'Quantity_Bags', 'qty': 'Quantity_Bags', 'bags': 'Quantity_Bags',
            'no_of_bags': 'Quantity_Bags', 'units': 'Quantity_Bags',
            
            # Rate columns
            'rate': 'Rate_Per_Bag', 'price': 'Rate_Per_Bag', 'unit_price': 'Rate_Per_Bag',
            'cost_per_bag': 'Rate_Per_Bag', 'rate_per_unit': 'Rate_Per_Bag',
            
            # Amount columns
            'amount': 'Total_Amount', 'total': 'Total_Amount', 'value': 'Total_Amount',
            'total_cost': 'Total_Amount', 'total_value': 'Total_Amount',
            
            # Supplier columns
            'supplier': 'Supplier', 'vendor': 'Supplier', 'dealer': 'Supplier',
            'seller': 'Supplier', 'party': 'Supplier',
            
            # Invoice columns
            'invoice': 'Invoice_Number', 'bill_no': 'Invoice_Number', 'receipt_no': 'Invoice_Number',
            'invoice_no': 'Invoice_Number', 'bill_number': 'Invoice_Number',
            
            # Vehicle columns
            'vehicle': 'Vehicle_Number', 'truck_no': 'Vehicle_Number', 'transport': 'Vehicle_Number',
            'vehicle_no': 'Vehicle_Number', 'lorry_no': 'Vehicle_Number',
            
            # Remarks columns
            'remarks': 'Remarks', 'notes': 'Remarks', 'comment': 'Remarks',
            'description': 'Remarks', 'memo': 'Remarks'
        }
        
        # Apply mappings (case insensitive)
        new_columns = []
        for col in df_clean.columns:
            mapped = False
            for key, value in column_mappings.items():
                if key.lower() in col.lower():
                    new_columns.append(value)
                    mapped = True
                    break
            if not mapped:
                new_columns.append(col)
        
        df_clean.columns = new_columns
        return df_clean
    
    def standardize_data(self, df: pd.DataFrame) -> pd.DataFrame:
        """Standardize data formats and add missing columns"""
        df_std = df.copy()
        
        # Ensure all standard columns exist
        for col in self.standard_columns:
            if col not in df_std.columns:
                df_std[col] = np.nan
        
        # Clean and format data
        try:
            # Date formatting
            if 'Date' in df_std.columns:
                df_std['Date'] = pd.to_datetime(df_std['Date'], errors='coerce')
            
            # Numeric formatting
            numeric_columns = ['Quantity_Bags', 'Rate_Per_Bag', 'Total_Amount']
            for col in numeric_columns:
                if col in df_std.columns:
                    # Remove currency symbols and convert to numeric
                    df_std[col] = df_std[col].astype(str).str.replace(r'[₹,\s]', '', regex=True)
                    df_std[col] = pd.to_numeric(df_std[col], errors='coerce')
            
            # Calculate Total_Amount if missing
            if df_std['Total_Amount'].isna().all() and not df_std['Quantity_Bags'].isna().all() and not df_std['Rate_Per_Bag'].isna().all():
                df_std['Total_Amount'] = df_std['Quantity_Bags'] * df_std['Rate_Per_Bag']
            
            # Clean text columns
            text_columns = ['Brand', 'Cement_Type', 'Supplier', 'Invoice_Number', 'Vehicle_Number', 'Remarks']
            for col in text_columns:
                if col in df_std.columns:
                    df_std[col] = df_std[col].astype(str).str.strip()
                    df_std[col] = df_std[col].replace('nan', np.nan)
        
        except Exception as e:
            st.warning(f"Some data standardization issues: {str(e)}")
        
        # Reorder columns
        column_order = [col for col in self.standard_columns if col in df_std.columns]
        other_columns = [col for col in df_std.columns if col not in self.standard_columns]
        df_std = df_std[column_order + other_columns]
        
        return df_std
    
    def add_summary_sheet(self, df: pd.DataFrame) -> Dict[str, pd.DataFrame]:
        """Create summary statistics"""
        summary_data = {}
        
        # Brand-wise summary
        if not df.empty:
            brand_summary = df.groupby('Brand').agg({
                'Quantity_Bags': 'sum',
                'Total_Amount': 'sum',
                'Date': ['min', 'max'],
                'Invoice_Number': 'count'
            }).round(2)
            
            brand_summary.columns = ['Total_Bags', 'Total_Amount', 'First_Date', 'Last_Date', 'Transaction_Count']
            summary_data['Brand_Summary'] = brand_summary.reset_index()
            
            # Monthly summary
            if 'Date' in df.columns and not df['Date'].isna().all():
                df_monthly = df.copy()
                df_monthly['Month_Year'] = df_monthly['Date'].dt.to_period('M').astype(str)
                monthly_summary = df_monthly.groupby('Month_Year').agg({
                    'Quantity_Bags': 'sum',
                    'Total_Amount': 'sum',
                    'Brand': 'nunique',
                    'Invoice_Number': 'count'
                }).round(2)
                
                monthly_summary.columns = ['Total_Bags', 'Total_Amount', 'Unique_Brands', 'Transaction_Count']
                summary_data['Monthly_Summary'] = monthly_summary.reset_index()
        
        return summary_data

def main():
    st.title("🏗️ Cement Ledger to Excel Converter")
    st.markdown("Upload your cement ledger files in various formats and convert them to standardized Excel files.")
    
    # Initialize processor
    processor = LedgerProcessor()
    
    # Sidebar for options
    st.sidebar.header("Options")
    include_summary = st.sidebar.checkbox("Include Summary Sheets", value=True)
    auto_calculate = st.sidebar.checkbox("Auto-calculate missing amounts", value=True)
    
    # File upload
    st.header("📁 Upload Ledger Files")
    uploaded_files = st.file_uploader(
        "Choose ledger files",
        accept_multiple_files=True,
        type=['csv', 'xlsx', 'xls', 'txt'],
        help="Upload CSV, Excel, or Text files containing cement ledger data"
    )
    
    if uploaded_files:
        # Process each file
        all_data = []
        file_summaries = []
        
        for file in uploaded_files:
            st.subheader(f"Processing: {file.name}")
            
            with st.spinner(f"Reading {file.name}..."):
                # Read file
                df_raw = processor.read_file(file)
                
                if df_raw.empty:
                    st.error(f"Could not read {file.name}")
                    continue
                
                # Show raw data preview
                with st.expander(f"Raw Data Preview - {file.name}"):
                    st.dataframe(df_raw.head(10))
                    st.info(f"Shape: {df_raw.shape[0]} rows, {df_raw.shape[1]} columns")
                
                # Clean and standardize
                df_clean = processor.clean_column_names(df_raw)
                df_standard = processor.standardize_data(df_clean)
                
                # Show processed data preview
                with st.expander(f"Processed Data Preview - {file.name}"):
                    st.dataframe(df_standard.head(10))
                    
                    # Show data quality info
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Records", len(df_standard))
                    with col2:
                        st.metric("Complete Records", len(df_standard.dropna()))
                    with col3:
                        total_amount = df_standard['Total_Amount'].sum() if 'Total_Amount' in df_standard.columns else 0
                        st.metric("Total Amount", f"₹{total_amount:,.2f}")
                
                # Add file identifier
                df_standard['Source_File'] = file.name
                all_data.append(df_standard)
                
                file_summaries.append({
                    'File': file.name,
                    'Records': len(df_standard),
                    'Brands': df_standard['Brand'].nunique() if 'Brand' in df_standard.columns else 0,
                    'Total_Amount': df_standard['Total_Amount'].sum() if 'Total_Amount' in df_standard.columns else 0
                })
        
        if all_data:
            # Combine all data
            combined_df = pd.concat(all_data, ignore_index=True, sort=False)
            
            # Show combined summary
            st.header("📊 Combined Data Summary")
            summary_df = pd.DataFrame(file_summaries)
            st.dataframe(summary_df)
            
            # Show overall metrics
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Files", len(uploaded_files))
            with col2:
                st.metric("Total Records", len(combined_df))
            with col3:
                st.metric("Unique Brands", combined_df['Brand'].nunique() if 'Brand' in combined_df.columns else 0)
            with col4:
                total_value = combined_df['Total_Amount'].sum() if 'Total_Amount' in combined_df.columns else 0
                st.metric("Total Value", f"₹{total_value:,.2f}")
            
            # Show final data preview
            st.header("📋 Final Standardized Data")
            st.dataframe(combined_df)
            
            # Prepare Excel file
            st.header("💾 Download Excel File")
            
            # Create Excel buffer
            excel_buffer = io.BytesIO()
            
            with pd.ExcelWriter(excel_buffer, engine='xlsxwriter') as writer:
                # Write main data
                combined_df.to_excel(writer, sheet_name='Cement_Ledger', index=False)
                
                # Add summary sheets if requested
                if include_summary:
                    summary_data = processor.add_summary_sheet(combined_df)
                    for sheet_name, summary_df in summary_data.items():
                        summary_df.to_excel(writer, sheet_name=sheet_name, index=False)
                
                # Add file summary
                summary_df.to_excel(writer, sheet_name='File_Summary', index=False)
                
                # Format the Excel file
                workbook = writer.book
                worksheet = writer.sheets['Cement_Ledger']
                
                # Add formatting
                header_format = workbook.add_format({
                    'bold': True,
                    'text_wrap': True,
                    'valign': 'top',
                    'fg_color': '#D7E4BC',
                    'border': 1
                })
                
                # Write headers with formatting
                for col_num, value in enumerate(combined_df.columns.values):
                    worksheet.write(0, col_num, value, header_format)
                
                # Auto-adjust column widths
                for column in combined_df:
                    column_width = max(combined_df[column].astype(str).map(len).max(), len(column))
                    col_idx = combined_df.columns.get_loc(column)
                    worksheet.set_column(col_idx, col_idx, min(column_width + 2, 50))
            
            excel_buffer.seek(0)
            
            # Generate filename
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"cement_ledger_converted_{timestamp}.xlsx"
            
            # Download button
            st.download_button(
                label="📥 Download Excel File",
                data=excel_buffer,
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                help="Click to download the standardized Excel file"
            )
            
            st.success(f"✅ Successfully processed {len(uploaded_files)} files with {len(combined_df)} total records!")
            
            # Show data quality report
            with st.expander("📈 Data Quality Report"):
                st.write("**Missing Data Analysis:**")
                missing_data = combined_df.isnull().sum()
                missing_percent = (missing_data / len(combined_df)) * 100
                quality_df = pd.DataFrame({
                    'Column': missing_data.index,
                    'Missing_Count': missing_data.values,
                    'Missing_Percentage': missing_percent.values
                }).round(2)
                st.dataframe(quality_df)
    
    else:
        # Show instructions
        st.info("👆 Please upload your ledger files to get started.")
        
        with st.expander("📖 Supported File Formats & Instructions"):
            st.markdown("""
            **Supported Formats:**
            - CSV files (.csv)
            - Excel files (.xlsx, .xls)
            - Text files (.txt) with delimited data
            
            **Expected Columns (any combination):**
            - Date, Brand, Cement Type, Quantity, Rate, Amount
            - Supplier, Invoice Number, Vehicle Number, Remarks
            
            **Features:**
            - Automatic column mapping and standardization
            - Data validation and cleaning
            - Summary statistics generation
            - Multiple file processing
            - Excel output with formatting
            
            **Tips:**
            - Files can have different column names - the app will try to map them automatically
            - Missing amounts will be calculated if quantity and rate are available
            - All data will be combined into a single standardized Excel file
            """)

if __name__ == "__main__":
    main()
