import streamlit as st
import pandas as pd
import openpyxl
from collections import OrderedDict
import plotly.express as px
from statsmodels.tsa.arima.model import ARIMA
from scipy import stats
import base64
from io import BytesIO
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan, acorr_ljungbox

def excel_editor_and_analyzer():
    st.header("Excel Editor and Analyzer")
    
    apply_custom_css()
    
    tab1, tab2 = st.tabs(["Excel Editor", "Data Analyzer"])
    
    with tab1:
        excel_editor()
    
    with tab2:
        data_analyzer()

def apply_custom_css():
    st.markdown("""
    <style>
        .stApp {
            background-color: #f0f2f6;
        }
        .excel-table {
            border-collapse: collapse;
            width: 100%;
            font-family: Arial, sans-serif;
        }
        .excel-table th, .excel-table td {
            border: 1px solid #b0b0b0;
            padding: 8px;
            text-align: left;
        }
        .excel-table tr:nth-child(even) {
            background-color: #f8f8f8;
        }
        .excel-table th {
            padding-top: 12px;
            padding-bottom: 12px;
            background-color: #4CAF50;
            color: white;
        }
        .stButton>button {
            background-color: #4CAF50;
            color: white;
            border: none;
            padding: 10px 24px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            margin: 4px 2px;
            cursor: pointer;
            border-radius: 4px;
        }
    </style>
    """, unsafe_allow_html=True)

def excel_editor():
    st.subheader("Excel Editor")
    
    uploaded_file = st.file_uploader("Choose an Excel file for editing", type="xlsx")
    
    if uploaded_file is not None:
        excel_file = openpyxl.load_workbook(uploaded_file)
        sheet = excel_file.active
        
        st.write("Original Excel Structure (First 5 Rows)")
        excel_html = create_excel_structure_html(sheet, max_rows=5)
        st.markdown(excel_html, unsafe_allow_html=True)
        
        merged_groups = get_merged_column_groups(sheet)
        df = create_dataframe_from_excel(uploaded_file, merged_groups)
        
        st.subheader("Select columns to delete")
        cols_to_delete = st.multiselect("Choose columns to remove", df.columns)
        if cols_to_delete:
            df = df.drop(columns=cols_to_delete)
            st.success(f"Deleted columns: {', '.join(cols_to_delete)}")
        
        st.subheader("Delete rows")
        num_rows = st.number_input("Enter the number of rows to delete from the start", min_value=0, max_value=len(df)-1, value=0)
        if num_rows > 0:
            df = df.iloc[num_rows:]
            st.success(f"Deleted first {num_rows} rows")
        
        st.subheader("Edit Data")
        edited_df = st.data_editor(df)
        
        st.markdown(get_excel_download_link(edited_df), unsafe_allow_html=True)
        
        if st.button("Upload Edited File to Home"):
            st.session_state.edited_df = edited_df
            st.session_state.edited_file_name = "edited_" + uploaded_file.name
            st.success("Edited file has been uploaded to Home. Please switch to the Home tab to see the uploaded file.")
    
    else:
        st.info("Please upload an Excel file to begin editing.")

def data_analyzer():
    st.subheader("Data Analyzer")
    
    uploaded_file = st.file_uploader("Choose an Excel file for analysis", type="xlsx")
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        st.write("Dataset Information:")
        st.write(f"Number of rows: {df.shape[0]}")
        st.write(f"Number of columns: {df.shape[1]}")
        
        numeric_columns = df.select_dtypes(include=['float64', 'int64']).columns
        x_col = st.selectbox("Select X-axis variable", numeric_columns)
        y_col = st.selectbox("Select Y-axis variable", numeric_columns)
        
        st.subheader("Advanced Visualization")
        chart_type = st.selectbox("Select chart type", ["Scatter", "Line", "Bar", "Box", "Violin", "3D Scatter"])
        
        if chart_type == "Scatter":
            fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        elif chart_type == "Line":
            fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        elif chart_type == "Bar":
            fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        elif chart_type == "Box":
            fig = px.box(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        elif chart_type == "Violin":
            fig = px.violin(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
        elif chart_type == "3D Scatter":
            z_col = st.selectbox("Select Z-axis variable", numeric_columns)
            fig = px.scatter_3d(df, x=x_col, y=y_col, z=z_col, title=f"3D Scatter Plot")
        
        st.plotly_chart(fig)
        
        st.subheader("Regression Analysis")
        regression_type = st.selectbox("Select regression type", ["Simple Linear", "Multiple Linear"])
        
        if regression_type == "Simple Linear":
            X = sm.add_constant(df[x_col])
            y = df[y_col]
            model = sm.OLS(y, X).fit()
            st.write(model.summary())
            
            fig = px.scatter(df, x=x_col, y=y_col, title=f"Simple Linear Regression: {y_col} vs {x_col}")
            fig.add_scatter(x=df[x_col], y=model.predict(X), mode='lines', name='Regression Line')
            st.plotly_chart(fig)
        
        elif regression_type == "Multiple Linear":
            independent_vars = st.multiselect("Select independent variables", numeric_columns, default=[x_col])
            if len(independent_vars) > 0:
                X = sm.add_constant(df[independent_vars])
                y = df[y_col]
                model = sm.OLS(y, X).fit()
                st.write(model.summary())
        
        st.subheader("Statistical Tests")
        
        st.write("Breusch-Pagan Test for Heteroscedasticity")
        _, p_value, _, _ = het_breuschpagan(model.resid, model.model.exog)
        st.write(f"p-value: {p_value:.4f}")
        st.write("Null hypothesis: Homoscedasticity")
        st.write(f"{'Reject' if p_value < 0.05 else 'Fail to reject'} the null hypothesis at 5% significance level.")
        
        st.write("Ljung-Box Test for Autocorrelation")
        lb_result = acorr_ljungbox(model.resid, lags=[1])
        st.write(f"p-value: {lb_result['lb_pvalue'][0]:.4f}")
        st.write("Null hypothesis: No autocorrelation")
        st.write(f"{'Reject' if lb_result['lb_pvalue'][0] < 0.05 else 'Fail to reject'} the null hypothesis at 5% significance level.")

    else:
        st.info("Please upload an Excel file to begin analysis.")

def create_excel_structure_html(sheet, max_rows=5):
    html = "<table class='excel-table'>"
    merged_cells = sheet.merged_cells.ranges

    for idx, row in enumerate(sheet.iter_rows(max_row=max_rows)):
        html += "<tr>"
        for cell in row:
            merged = False
            for merged_range in merged_cells:
                if cell.coordinate in merged_range:
                    if cell.coordinate == merged_range.start_cell.coordinate:
                        rowspan = min(merged_range.max_row - merged_range.min_row + 1, max_rows - idx)
                        colspan = merged_range.max_col - merged_range.min_col + 1
                        html += f"<td rowspan='{rowspan}' colspan='{colspan}'>{cell.value}</td>"
                    merged = True
                    break
            if not merged:
                html += f"<td>{cell.value}</td>"
        html += "</tr>"
    html += "</table>"
    return html

def get_merged_column_groups(sheet):
    merged_groups = {}
    for merged_range in sheet.merged_cells.ranges:
        if merged_range.min_row == 1:  # Only consider merged cells in the first row (header)
            main_col = sheet.cell(1, merged_range.min_col).value
            merged_groups[main_col] = list(range(merged_range.min_col, merged_range.max_col + 1))
    return merged_groups

def create_dataframe_from_excel(file, merged_groups):
    df = pd.read_excel(file, header=None)
    column_headers = []
    column_indices = OrderedDict()
    
    for col in range(df.shape[1]):
        cell_value = df.iloc[0, col]
        if pd.notna(cell_value):
            column_headers.append(cell_value)
            if cell_value not in column_indices:
                column_indices[cell_value] = []
            column_indices[cell_value].append(col)
        else:
            prev_header = column_headers[-1]
            column_headers.append(prev_header)
            column_indices[prev_header].append(col)
    
    df.columns = column_headers
    df = df.iloc[1:]  # Remove the header row
    return df

def get_excel_download_link(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    excel_data = output.getvalue()
    b64 = base64.b64encode(excel_data).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="edited_file.xlsx">Download Edited Excel File</a>'

if __name__ == "__main__":
    excel_editor_and_analyzer()
