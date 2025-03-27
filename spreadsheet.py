import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objs as go

class SpreadsheetApp:
    def __init__(self):
        # Initialize session state for spreadsheet data
        if 'spreadsheet_data' not in st.session_state:
            st.session_state.spreadsheet_data = pd.DataFrame()
        
        # Initialize session state for additional features
        if 'formatting' not in st.session_state:
            st.session_state.formatting = {}
        
        # Initialize session state for charts
        if 'charts' not in st.session_state:
            st.session_state.charts = []

    def create_new_sheet(self):
        """Create a new blank spreadsheet"""
        rows = st.number_input("Number of Rows", min_value=1, max_value=100, value=10)
        cols = st.number_input("Number of Columns", min_value=1, max_value=26, value=5)
        
        # Create a blank DataFrame with alphabetic column names
        col_names = [chr(65 + i) for i in range(cols)]
        st.session_state.spreadsheet_data = pd.DataFrame(
            np.nan, 
            index=range(rows), 
            columns=col_names
        )
        st.success("New spreadsheet created!")

    def load_file(self):
        """Load an existing Excel or CSV file"""
        uploaded_file = st.file_uploader("Choose a file", type=['xlsx', 'csv'])
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.xlsx'):
                    st.session_state.spreadsheet_data = pd.read_excel(uploaded_file)
                else:
                    st.session_state.spreadsheet_data = pd.read_csv(uploaded_file)
                st.success("File loaded successfully!")
            except Exception as e:
                st.error(f"Error loading file: {e}")

    def edit_sheet(self):
        """Edit the spreadsheet with Streamlit data editor"""
        if st.session_state.spreadsheet_data is not None and not st.session_state.spreadsheet_data.empty:
            edited_df = st.data_editor(
                st.session_state.spreadsheet_data, 
                num_rows="dynamic",
                column_config={
                    col: st.column_config.TextColumn(col) 
                    for col in st.session_state.spreadsheet_data.columns
                }
            )
            st.session_state.spreadsheet_data = edited_df

    def data_analysis_tools(self):
        """Provide data analysis and manipulation tools"""
        st.subheader("Data Analysis Tools")
        
        # Basic statistical analysis
        if st.button("Show Basic Statistics"):
            st.write(st.session_state.spreadsheet_data.describe())
        
        # Data filtering
        st.subheader("Filter Data")
        filter_col = st.selectbox("Select Column to Filter", 
                                  st.session_state.spreadsheet_data.columns)
        filter_type = st.selectbox("Filter Type", 
                                   ["Equal to", "Greater than", "Less than"])
        
        if filter_type == "Equal to":
            filter_value = st.text_input("Value")
            if st.button("Apply Filter"):
                filtered_df = st.session_state.spreadsheet_data[
                    st.session_state.spreadsheet_data[filter_col] == filter_value
                ]
                st.write(filtered_df)
        
        # Basic charting
        st.subheader("Quick Visualization")
        x_col = st.selectbox("X-axis", st.session_state.spreadsheet_data.columns)
        y_col = st.selectbox("Y-axis", st.session_state.spreadsheet_data.columns)
        
        if st.button("Create Line Chart"):
            fig = px.line(st.session_state.spreadsheet_data, x=x_col, y=y_col)
            st.plotly_chart(fig)

    def export_file(self):
        """Export the current spreadsheet"""
        if st.session_state.spreadsheet_data is not None and not st.session_state.spreadsheet_data.empty:
            export_format = st.selectbox("Select Export Format", 
                                         ["Excel (.xlsx)", "CSV (.csv)"])
            
            if st.button("Download File"):
                if export_format == "Excel (.xlsx)":
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        st.session_state.spreadsheet_data.to_excel(writer, index=False)
                    excel_data = output.getvalue()
                    st.download_button(
                        label="Download Excel File",
                        data=excel_data,
                        file_name="spreadsheet_export.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    csv_data = st.session_state.spreadsheet_data.to_csv(index=False)
                    st.download_button(
                        label="Download CSV File",
                        data=csv_data,
                        file_name="spreadsheet_export.csv",
                        mime="text/csv"
                    )

    def advanced_formulas(self):
        """Implement basic spreadsheet formulas"""
        st.subheader("Formula Operations")
        
        # Column selection for formula
        col1 = st.selectbox("Select First Column", 
                            st.session_state.spreadsheet_data.columns)
        col2 = st.selectbox("Select Second Column", 
                            st.session_state.spreadsheet_data.columns)
        
        formula_type = st.selectbox("Formula Type", [
            "Sum", "Average", "Max", "Min", 
            "Add Columns", "Subtract Columns"
        ])
        
        if st.button("Apply Formula"):
            try:
                if formula_type == "Sum":
                    result = st.session_state.spreadsheet_data[col1].sum()
                elif formula_type == "Average":
                    result = st.session_state.spreadsheet_data[col1].mean()
                elif formula_type == "Max":
                    result = st.session_state.spreadsheet_data[col1].max()
                elif formula_type == "Min":
                    result = st.session_state.spreadsheet_data[col1].min()
                elif formula_type == "Add Columns":
                    result = st.session_state.spreadsheet_data[col1] + st.session_state.spreadsheet_data[col2]
                elif formula_type == "Subtract Columns":
                    result = st.session_state.spreadsheet_data[col1] - st.session_state.spreadsheet_data[col2]
                
                st.write("Result:", result)
            except Exception as e:
                st.error(f"Error applying formula: {e}")

def main():
    # App title and description
    st.title("📊 Excel-Like Spreadsheet Application")
    st.write("A powerful spreadsheet tool with multiple features")

    # Initialize the SpreadsheetApp
    app = SpreadsheetApp()

    # Sidebar for navigation
    menu = st.sidebar.radio("Menu", [
        "Create/Load Sheet", 
        "Edit Sheet", 
        "Data Analysis", 
        "Advanced Formulas", 
        "Export Sheet"
    ])

    # Navigation logic
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

    # Display current sheet if it exists
    if not st.session_state.spreadsheet_data.empty:
        st.subheader("Current Spreadsheet")
        st.dataframe(st.session_state.spreadsheet_data)

if __name__ == "__main__":
    main()
