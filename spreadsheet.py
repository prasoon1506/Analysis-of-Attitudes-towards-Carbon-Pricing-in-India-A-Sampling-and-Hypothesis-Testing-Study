import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objs as go
class SpreadsheetApp:
    def __init__(self):
        if 'spreadsheet_data' not in st.session_state:
            st.session_state.spreadsheet_data = pd.DataFrame()
        if 'formatting' not in st.session_state:
            st.session_state.formatting = {}
        if 'charts' not in st.session_state:
            st.session_state.charts = []
    def create_new_sheet(self):
        rows = st.number_input("Number of Rows", min_value=1, max_value=100, value=10)
        cols = st.number_input("Number of Columns", min_value=1, max_value=26, value=5)
        col_names = [chr(65 + i) for i in range(cols)]
        st.session_state.spreadsheet_data = pd.DataFrame('',index=range(rows),columns=col_names)
        st.success("New spreadsheet created!")
    def load_file(self):
        uploaded_file = st.file_uploader("Choose a file", type=['xlsx', 'csv'])
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.xlsx'):
                    df = pd.read_excel(uploaded_file)
                else:
                    df = pd.read_csv(uploaded_file)
                df = df.astype(str)
                st.session_state.spreadsheet_data = df
                st.success("File loaded successfully!")
            except Exception as e:
                st.error(f"Error loading file: {e}")
    def edit_sheet(self):
        if st.session_state.spreadsheet_data is not None and not st.session_state.spreadsheet_data.empty:
            column_config = {col: st.column_config.TextColumn(col, required=False) for col in st.session_state.spreadsheet_data.columns}
            edited_df = st.data_editor(st.session_state.spreadsheet_data,num_rows="dynamic",column_config=column_config,use_container_width=True)
            st.session_state.spreadsheet_data = edited_df
    def data_analysis_tools(self):
        st.subheader("Data Analysis Tools")
        try:
            numeric_df = st.session_state.spreadsheet_data.apply(pd.to_numeric, errors='ignore')
        except:
            numeric_df = st.session_state.spreadsheet_data
        if st.button("Show Basic Statistics"):
            st.write(numeric_df.describe())
        st.subheader("Filter Data")
        filter_col = st.selectbox("Select Column to Filter", st.session_state.spreadsheet_data.columns)
        filter_type = st.selectbox("Filter Type", ["Equal to", "Contains"])
        if filter_type == "Equal to":
            filter_value = st.text_input("Value")
            if st.button("Apply Filter"):
                filtered_df = st.session_state.spreadsheet_data[st.session_state.spreadsheet_data[filter_col] == filter_value]
                st.write(filtered_df)
        elif filter_type == "Contains":
            filter_value = st.text_input("Value")
            if st.button("Apply Contains Filter"):
                filtered_df = st.session_state.spreadsheet_data[st.session_state.spreadsheet_data[filter_col].str.contains(filter_value, case=False, na=False)]
                st.write(filtered_df)
        st.subheader("Quick Visualization")
        numeric_columns = numeric_df.select_dtypes(include=[np.number]).columns.tolist()
        if numeric_columns:
            x_col = st.selectbox("X-axis", numeric_columns)
            y_col = st.selectbox("Y-axis", numeric_columns)
            if st.button("Create Line Chart"):
                try:
                    fig = px.line(numeric_df, x=x_col, y=y_col)
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error creating chart: {e}")
        else:
            st.warning("No numeric columns available for charting")
    def export_file(self):
        if st.session_state.spreadsheet_data is not None and not st.session_state.spreadsheet_data.empty:
            export_format = st.selectbox("Select Export Format", ["Excel (.xlsx)", "CSV (.csv)"])
            if st.button("Download File"):
                if export_format == "Excel (.xlsx)":
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        st.session_state.spreadsheet_data.to_excel(writer, index=False)
                    excel_data = output.getvalue()
                    st.download_button(label="Download Excel File",data=excel_data,file_name="spreadsheet_export.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                else:
                    csv_data = st.session_state.spreadsheet_data.to_csv(index=False)
                    st.download_button(label="Download CSV File",data=csv_data,file_name="spreadsheet_export.csv",mime="text/csv")
    def advanced_formulas(self):
        st.subheader("Formula Operations")
        try:
            numeric_df = st.session_state.spreadsheet_data.apply(pd.to_numeric, errors='coerce')
        except:
            st.error("Unable to perform numeric operations")
            return
        numeric_columns = numeric_df.select_dtypes(include=[np.number]).columns.tolist()
        if not numeric_columns:
            st.warning("No numeric columns available for formula operations")
            return
        col1 = st.selectbox("Select First Column", numeric_columns)
        col2 = st.selectbox("Select Second Column", numeric_columns)
        formula_type = st.selectbox("Formula Type", ["Sum", "Average", "Max", "Min","Add Columns", "Subtract Columns"])
        if st.button("Apply Formula"):
            try:
                if formula_type == "Sum":
                    result = numeric_df[col1].sum()
                elif formula_type == "Average":
                    result = numeric_df[col1].mean()
                elif formula_type == "Max":
                    result = numeric_df[col1].max()
                elif formula_type == "Min":
                    result = numeric_df[col1].min()
                elif formula_type == "Add Columns":
                    result = numeric_df[col1] + numeric_df[col2]
                elif formula_type == "Subtract Columns":
                    result = numeric_df[col1] - numeric_df[col2]
                st.write("Result:", result)
            except Exception as e:
                st.error(f"Error applying formula: {e}")
def main():
    st.title("📊 Excel-Like Spreadsheet Application")
    st.write("A powerful spreadsheet tool with multiple features")
    app = SpreadsheetApp()
    menu = st.sidebar.radio("Menu", ["Create/Load Sheet","Edit Sheet","Data Analysis","Advanced Formulas","Export Sheet"])
    if menu == "Create/Load Sheet":
        st.header("Create or Load Spreadsheet")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Create New Sheet"):
                app.create_new_sheet()
        with col2:
            app.load_file()
    elif menu == "Edit Sheet":
        st.header("Edit Spreadsheet")
        app.edit_sheet()
    elif menu == "Data Analysis":
        st.header("Data Analysis Tools")
        app.data_analysis_tools()
    elif menu == "Advanced Formulas":
        st.header("Advanced Formula Operations")
        app.advanced_formulas()
    elif menu == "Export Sheet":
        st.header("Export Spreadsheet")
        app.export_file()
    if not st.session_state.spreadsheet_data.empty:
        st.subheader("Current Spreadsheet")
        st.dataframe(st.session_state.spreadsheet_data)
if __name__ == "__main__":
    main()
