import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xgboost as xgb
from sklearn.metrics import mean_squared_error
from scipy import stats
from io import BytesIO
import base64
import matplotlib.backends.backend_pdf

# Set page config
st.set_page_config(page_title="Brand Price Analysis", layout="wide")

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'headers' not in st.session_state:
    st.session_state.headers = None
if 'district_benchmarks' not in st.session_state:
    st.session_state.district_benchmarks = {}

def read_headers(file):
    df_headers = pd.read_excel(file, nrows=2)
    headers = df_headers.iloc[0].tolist()
    sub_headers = df_headers.iloc[1].tolist()
    
    week_headers = []
    for header, sub_header in zip(headers, sub_headers):
        if pd.notna(header) and 'GAP' not in str(header):
            week_headers.append(str(header))
        elif pd.notna(sub_header):
            week_headers.append(sub_header)
    
    return week_headers

def transform_data(df, selected_weeks):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    for week in selected_weeks:
        week_data = df[[f"{week} {brand}" for brand in brands]]
        week_data = week_data.rename(columns={f"{week} {brand}": f"{brand} ({week})" for brand in brands})
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True)

    return transformed_df

def plot_district_graph(df, district_name, benchmark_brands, desired_diff):
    # ... (rest of the function remains the same)

def generate_pdf(figs):
    # ... (function remains the same)

def main():
    st.title("Brand Price Analysis")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            st.session_state.headers = read_headers(uploaded_file)
            df = pd.read_excel(uploaded_file, skiprows=2)
            st.success("File uploaded successfully!")
            
            # Add dropdown for selecting weeks/months
            selected_weeks = st.multiselect("Select Weeks/Months for Analysis", options=st.session_state.headers)
            
            if selected_weeks:
                st.session_state.df = transform_data(df, selected_weeks)
            else:
                st.warning("Please select at least one week/month for analysis.")
                return
        except Exception as e:
            st.error(f"Error reading file: {e}. Please ensure it is a valid Excel file.")
            return

    if st.session_state.df is not None:
        df = st.session_state.df
        
        # ... (rest of the main function remains the same)

if __name__ == "__main__":
    main()
