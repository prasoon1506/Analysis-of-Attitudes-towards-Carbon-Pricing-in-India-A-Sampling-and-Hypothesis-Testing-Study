import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
import io
import re
from typing import Dict, List, Optional, Union

# PDF processing imports with error handling
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False

try:
    import PyPDF2
    PYPDF2_AVAILABLE = True
except ImportError:
    PYPDF2_AVAILABLE = False

try:
    import tabula
    TABULA_AVAILABLE = True
except ImportError:
    TABULA_AVAILABLE = False

try:
    import fitz  # PyMuPDF
    FITZ_AVAILABLE = True
except ImportError:
    FITZ_AVAILABLE = False

try:
    from PIL import Image
    import pytesseract
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False

# Configure Streamlit page
st.set_page_config(
    page_title="Cement Ledger to Excel Converter",
    page_icon="🏗️",
    layout="wide"
)

# Check for required packages
@st.cache_data
def check_pdf_dependencies():
    """Check if PDF processing libraries are available"""
    available = {
        'pdfplumber': PDFPLUMBER_AVAILABLE,
        'PyPDF2': PYPDF2_AVAILABLE,
        'tabula-py': TABULA_AVAILABLE,
        'PyMuPDF': FITZ_AVAILABLE,
        'OCR (pytesseract)': OCR_AVAILABLE
    }
    
    missing = [name for name, avail in available.items() if not avail]
    return available, missing

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
        elif file.name.endswith('.pdf'):
            return 'pdf'
        else:
            return 'unknown'
    
    def extract_text_from_pdf(self, file) -> str:
        """Extract text from PDF using multiple methods"""
        text = ""
        file.seek(0)
        
        # Method 1: Try pdfplumber (best for tables)
        if PDFPLUMBER_AVAILABLE:
            try:
                with pdfplumber.open(file) as pdf:
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text + "\n"
                    if text.strip():
                        st.success("✅ PDFPlumber text extraction successful")
                        return text
            except Exception as e:
                st.warning(f"PDFPlumber text extraction failed: {str(e)}")
        
        # Method 2: Try PyPDF2
        if PYPDF2_AVAILABLE:
            file.seek(0)
            try:
                reader = PyPDF2.PdfReader(file)
                for page in reader.pages:
                    page_text = page.extract_text()
                    if page_text:
                        text += page_text + "\n"
                if text.strip():
                    st.success("✅ PyPDF2 text extraction successful")
                    return text
            except Exception as e:
                st.warning(f"PyPDF2 failed: {str(e)}")
        
        # Method 3: Try PyMuPDF (fitz)
        if FITZ_AVAILABLE:
            file.seek(0)
            try:
                pdf_document = fitz.open(stream=file.read(), filetype="pdf")
                for page_num in range(pdf_document.page_count):
                    page = pdf_document[page_num]
                    page_text = page.get_text()
                    if page_text:
                        text += page_text + "\n"
                pdf_document.close()
                if text.strip():
                    st.success("✅ PyMuPDF text extraction successful")
                    return text
            except Exception as e:
                st.warning(f"PyMuPDF failed: {str(e)}")
        
        return text

    def extract_tables_from_pdf(self, file) -> List[pd.DataFrame]:
        """Extract tables from PDF using multiple methods"""
        tables = []
        file.seek(0)
        
        # Method 1: Try tabula-py (best for structured tables)
        if TABULA_AVAILABLE:
            try:
                file.seek(0)
                # Create temporary file for tabula
                import tempfile
                import os
                
                with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as temp_file:
                    temp_file.write(file.read())
                    temp_file_path = temp_file.name
                
                # Extract tables using tabula
                tabula_tables = tabula.read_pdf(
                    temp_file_path, 
                    pages='all', 
                    multiple_tables=True,
                    pandas_options={'header': 0}
                )
                
                # Clean up temporary file
                os.unlink(temp_file_path)
                
                for table in tabula_tables:
                    if not table.empty and table.shape[0] > 1:
                        tables.append(table)
                
                if tables:
                    st.success(f"✅ Tabula extracted {len(tables)} tables")
                    return tables
                    
            except Exception as e:
                st.warning(f"Tabula-py extraction failed: {str(e)}")
        
        # Method 2: Try pdfplumber for table extraction
        if PDFPLUMBER_AVAILABLE:
            file.seek(0)
            try:
                with pdfplumber.open(file) as pdf:
                    for page_num, page in enumerate(pdf.pages):
                        page_tables = page.extract_tables()
                        for table_num, table in enumerate(page_tables):
                            if table and len(table) > 1:
                                # Convert to DataFrame
                                try:
                                    df = pd.DataFrame(table[1:], columns=table[0])
                                    # Clean empty columns and rows
                                    df = df.dropna(how='all').dropna(axis=1, how='all')
                                    if not df.empty:
                                        tables.append(df)
                                except Exception as e:
                                    st.warning(f"Error processing table {table_num+1} from page {page_num+1}: {str(e)}")
                
                if tables:
                    st.success(f"✅ PDFPlumber extracted {len(tables)} tables")
                    return tables
                    
            except Exception as e:
                st.warning(f"PDFPlumber table extraction failed: {str(e)}")
        
        return tables

    def parse_text_to_dataframe(self, text: str) -> pd.DataFrame:
        """Parse extracted text into a structured DataFrame"""
        if not text.strip():
            return pd.DataFrame()
        
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        # Try to identify tabular data patterns
        potential_data = []
        headers = None
        
        for line in lines:
            # Skip lines that look like headers/titles
            if any(keyword in line.lower() for keyword in ['ledger', 'statement', 'report', 'cement', 'company']):
                continue
            
            # Look for lines with multiple data points separated by spaces/tabs
            parts = re.split(r'\s{2,}|\t', line)  # Split on multiple spaces or tabs
            if len(parts) >= 3:  # At least 3 columns for meaningful data
                if not headers and any(keyword in line.lower() for keyword in ['date', 'brand', 'quantity', 'amount', 'rate']):
                    headers = [part.strip() for part in parts]
                else:
                    potential_data.append([part.strip() for part in parts])
        
        # If no clear headers found, create generic ones
        if not headers and potential_data:
            max_cols = max(len(row) for row in potential_data)
            headers = [f'Column_{i+1}' for i in range(max_cols)]
        
        # Create DataFrame
        if headers and potential_data:
            # Ensure all rows have same number of columns
            max_cols = len(headers)
            cleaned_data = []
            for row in potential_data:
                while len(row) < max_cols:
                    row.append('')
                cleaned_data.append(row[:max_cols])
            
            return pd.DataFrame(cleaned_data, columns=headers)
        
        # If structured parsing fails, try line-by-line parsing
        return self.parse_unstructured_text(text)

    def parse_unstructured_text(self, text: str) -> pd.DataFrame:
        """Parse unstructured text by looking for patterns"""
        lines = [line.strip() for line in text.split('\n') if line.strip()]
        
        data_rows = []
        for line in lines:
            # Look for date patterns
            date_match = re.search(r'\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b', line)
            
            # Look for amount patterns
            amount_matches = re.findall(r'₹?\s*\d+(?:,\d{3})*(?:\.\d{2})?', line)
            
            # Look for quantity patterns
            qty_match = re.search(r'\b(\d+)\s*(?:bags?|units?|nos?)\b', line, re.IGNORECASE)
            
            if date_match or amount_matches or qty_match:
                row_data = {
                    'Raw_Text': line,
                    'Date': date_match.group() if date_match else '',
                    'Amounts': ', '.join(amount_matches) if amount_matches else '',
                    'Quantity': qty_match.group(1) if qty_match else ''
                }
                
                # Extract brand/supplier info (words in title case)
                words = line.split()
                title_case_words = [word for word in words if word.istitle() and len(word) > 2]
                row_data['Potential_Brand'] = ' '.join(title_case_words[:2])  # Take first 2 title case words
                
                data_rows.append(row_data)
        
        return pd.DataFrame(data_rows) if data_rows else pd.DataFrame({'Raw_Data': lines})
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
            
            elif file_type == 'pdf':
                return self.process_pdf(file)
            
            else:
                st.error(f"Unsupported file type: {file.name}")
                return pd.DataFrame()
                
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            return pd.DataFrame()
    
    def process_pdf(self, file) -> pd.DataFrame:
        """Process PDF file using multiple extraction methods"""
        st.info(f"🔍 Processing PDF: {file.name}")
        
        # Check if any PDF libraries are available
        if not any([PDFPLUMBER_AVAILABLE, PYPDF2_AVAILABLE, TABULA_AVAILABLE, FITZ_AVAILABLE]):
            st.error("❌ No PDF processing libraries available. Please install: pip install pdfplumber PyPDF2 tabula-py PyMuPDF")
            return pd.DataFrame({'Error': ['No PDF processing libraries available']})
        
        # Try table extraction first
        with st.spinner("Extracting tables from PDF..."):
            tables = self.extract_tables_from_pdf(file)
            
            if tables:
                st.success(f"✅ Found {len(tables)} tables in PDF")
                
                # If multiple tables, combine them or let user choose
                if len(tables) == 1:
                    return tables[0]
                else:
                    # Show preview of tables and combine
                    st.write("**Found multiple tables:**")
                    combined_df = pd.DataFrame()
                    
                    for i, table in enumerate(tables):
                        with st.expander(f"Table {i+1} Preview"):
                            st.dataframe(table.head(3))
                        
                        # Combine tables with similar structures
                        if combined_df.empty:
                            combined_df = table
                        else:
                            try:
                                combined_df = pd.concat([combined_df, table], ignore_index=True, sort=False)
                            except:
                                # If can't combine, add as separate columns
                                table.columns = [f"Table{i+1}_{col}" for col in table.columns]
                                combined_df = pd.concat([combined_df, table], axis=1)
                    
                    return combined_df
        
        # If no tables found, try text extraction
        with st.spinner("Extracting text from PDF..."):
            text = self.extract_text_from_pdf(file)
            
            if text.strip():
                st.info("📄 No structured tables found. Parsing text data...")
                return self.parse_text_to_dataframe(text)
            else:
                st.warning("⚠️ Could not extract readable text from PDF.")
                
                # Try OCR if enabled and available
                if OCR_AVAILABLE:
                    st.info("🔍 Attempting OCR extraction...")
                    ocr_text = self.ocr_pdf_page(file)
                    if ocr_text.strip():
                        return self.parse_text_to_dataframe(ocr_text)
                
                st.error("❌ Could not extract data from PDF. The PDF might be:")
                st.write("- Scanned/image-based (try enabling OCR)")
                st.write("- Password protected")
                st.write("- Corrupted or in an unsupported format")
                st.write("💡 Try converting the PDF to Excel/CSV manually first")
                
                return pd.DataFrame({'Error': ['Could not extract data from PDF']})
    
    def ocr_pdf_page(self, file) -> str:
        """Extract text from PDF using OCR (for scanned PDFs)"""
        if not OCR_AVAILABLE:
            st.warning("OCR libraries not available. Install: pip install pytesseract pillow")
            return ""
        
        try:
            # Convert PDF pages to images and apply OCR
            if FITZ_AVAILABLE:
                file.seek(0)
                pdf_document = fitz.open(stream=file.read(), filetype="pdf")
                text = ""
                
                for page_num in range(min(pdf_document.page_count, 3)):  # Limit to first 3 pages
                    page = pdf_document[page_num]
                    pix = page.get_pixmap()
                    img_data = pix.tobytes("png")
                    
                    # Apply OCR
                    image = Image.open(io.BytesIO(img_data))
                    page_text = pytesseract.image_to_string(image)
                    text += page_text + "\n"
                
                pdf_document.close()
                
                if text.strip():
                    st.success("✅ OCR extraction successful")
                    return text
            else:
                st.warning("PyMuPDF not available for OCR. Install: pip install PyMuPDF")
                
        except Exception as e:
            st.warning(f"OCR failed: {str(e)}")
        
        return ""
    
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
    st.markdown("Upload your cement ledger files in various formats (CSV, Excel, Text, **PDF**) and convert them to standardized Excel files.")
    
    # Check PDF dependencies
    available_libs, missing_deps = check_pdf_dependencies()
    
    # Show available libraries status
    if any(available_libs.values()):
        st.sidebar.success("📚 Available PDF Libraries:")
        for lib, status in available_libs.items():
            if status:
                st.sidebar.success(f"✅ {lib}")
            else:
                st.sidebar.error(f"❌ {lib}")
    
    if missing_deps:
        st.warning(f"⚠️ For enhanced PDF support, install missing libraries")
        with st.expander("📦 Installation Instructions"):
            st.code(f"""
# Install missing PDF processing libraries
pip install {' '.join(missing_deps).replace('tabula-py', 'tabula-py').replace('OCR (pytesseract)', 'pytesseract pillow')}

# For OCR support (optional - for scanned PDFs):
pip install pytesseract pillow

# Install tesseract OCR engine:
# Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
# Mac: brew install tesseract
# Linux: sudo apt-get install tesseract-ocr
            """)
        
        # Show current capabilities
        st.info(f"✅ Currently available: {', '.join([lib for lib, avail in available_libs.items() if avail]) or 'Basic file processing only'}")
    
    # Initialize processor
    processor = LedgerProcessor()
    
    # Sidebar for options
    st.sidebar.header("Options")
    include_summary = st.sidebar.checkbox("Include Summary Sheets", value=True)
    auto_calculate = st.sidebar.checkbox("Auto-calculate missing amounts", value=True)
    
    # PDF processing options
    st.sidebar.subheader("PDF Processing Options")
    pdf_method = st.sidebar.selectbox(
        "PDF Extraction Method",
        ["Auto-detect", "Tables First", "Text Only"],
        help="Choose how to process PDF files"
    )
    ocr_enabled = st.sidebar.checkbox("Enable OCR for scanned PDFs", value=False, 
                                     help="Requires tesseract installation")
    
    # File upload
    st.header("📁 Upload Ledger Files")
    uploaded_files = st.file_uploader(
        "Choose ledger files",
        accept_multiple_files=True,
        type=['csv', 'xlsx', 'xls', 'txt', 'pdf'],
        help="Upload CSV, Excel, Text, or PDF files containing cement ledger data"
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
            - **PDF files (.pdf)** - Tables and text-based ledgers
            
            **PDF Processing Features:**
            - Automatic table extraction from structured PDFs
            - Text parsing for unstructured PDFs
            - Multiple extraction methods (pdfplumber, tabula, PyPDF2)
            - OCR support for scanned PDFs (requires setup)
            
            **Expected Columns (any combination):**
            - Date, Brand, Cement Type, Quantity, Rate, Amount
            - Supplier, Invoice Number, Vehicle Number, Remarks
            
            **Features:**
            - Automatic column mapping and standardization
            - Data validation and cleaning
            - Summary statistics generation
            - Multiple file processing
            - Excel output with formatting
            
            **PDF Tips:**
            - Works best with text-based PDFs containing tables
            - For scanned PDFs, enable OCR option (requires tesseract)
            - Multiple tables in a PDF will be automatically combined
            - If extraction fails, try converting PDF to Excel/CSV first
            
            **Installation for full PDF support:**
            ```bash
            pip install pdfplumber PyPDF2 tabula-py PyMuPDF
            # For OCR support:
            pip install pytesseract pillow
            # Install tesseract: https://github.com/tesseract-ocr/tesseract
            ```
            """)
        
        st.markdown("""
        ### 📋 PDF Processing Status
        The app will try multiple methods to extract data from your PDFs:
        1. **Table Extraction** - Best for structured ledgers with clear tables
        2. **Text Parsing** - For text-based PDFs without clear table structure  
        3. **OCR (Optional)** - For scanned/image-based PDFs
        
        Upload your PDF files and the app will automatically choose the best extraction method!
        """)

if __name__ == "__main__":
    main()
