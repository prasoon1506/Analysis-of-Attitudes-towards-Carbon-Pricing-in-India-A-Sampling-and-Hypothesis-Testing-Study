import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta
import plotly.graph_objects as go
from statistics import mode, median

st.set_page_config(layout="wide", page_title="Dealer Price Analysis Dashboard")

def load_data():
    """Load the uploaded dataset and perform initial preprocessing"""
    st.sidebar.title("Data Upload")
    uploaded_file = st.sidebar.file_uploader("Upload your dataset (CSV or Excel)", type=["csv", "xlsx", "xls"])
    
    if uploaded_file is not None:
        try:
            # Determine file type and read accordingly
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            
            # Check if required columns exist
            required_columns = [
                'District: Name', 'checkin date', 'Product Type', 'Brand: Name',
                'Account: Dealer Category', 'Whole Sale Price', 'Owner: Full Name'
            ]
            
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                st.error(f"Missing required columns: {', '.join(missing_columns)}")
                return None
            
            # Convert date column to datetime
            try:
                df['checkin date'] = pd.to_datetime(df['checkin date'])
            except Exception as e:
                st.error(f"Error converting 'checkin date' to datetime: {e}")
                return None
            
            # Fill NaN in dealer category
            df['Account: Dealer Category'] = df['Account: Dealer Category'].fillna('NaN')
            
            # Ensure price column is numeric
            df['Whole Sale Price'] = pd.to_numeric(df['Whole Sale Price'], errors='coerce')
            
            return df
        
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return None
    
    return None

def create_filters(df):
    """Create filter dropdowns based on the dataset"""
    st.sidebar.header("Filters")
    
    # Make a copy of the dataframe to avoid warnings
    df = df.copy()
    
    # District filter
    district_values = df['District: Name'].fillna('Unknown').unique().tolist()
    district_options = ['All'] + sorted(district_values)
    selected_district = st.sidebar.selectbox("Select District", district_options)
    
    # Date range filter
    min_date = df['checkin date'].min().date()
    max_date = df['checkin date'].max().date()
    
    date_selection_type = st.sidebar.radio("Date Selection", ["Single Date", "Date Range"])
    
    if date_selection_type == "Single Date":
        selected_date = st.sidebar.date_input("Select Checkin Date", min_date, min_value=min_date, max_value=max_date)
        date_filter = (df['checkin date'].dt.date == selected_date)
    else:
        date_range = st.sidebar.date_input("Select Date Range", 
                                          [min_date, min_date + timedelta(days=7)],
                                          min_value=min_date, 
                                          max_value=max_date)
        
        if len(date_range) == 2:
            start_date, end_date = date_range
            date_filter = (df['checkin date'].dt.date >= start_date) & (df['checkin date'].dt.date <= end_date)
        else:
            st.sidebar.warning("Please select both start and end dates")
            date_filter = pd.Series(True, index=df.index)
    
    # Product Type filter
    product_values = df['Product Type'].fillna('Unknown').unique().tolist()
    product_options = ['All'] + sorted(product_values)
    selected_product = st.sidebar.selectbox("Select Product Type", product_options)
    
    # Brand filter
    brand_values = df['Brand: Name'].fillna('Unknown').unique().tolist()
    brand_options = ['All'] + sorted(brand_values)
    selected_brand = st.sidebar.selectbox("Select Brand", brand_options)
    
    # Apply filters
    filtered_df = df.copy()
    
    if selected_district != 'All':
        filtered_df = filtered_df[filtered_df['District: Name'] == selected_district]
        
    filtered_df = filtered_df[date_filter]
    
    if selected_product != 'All':
        filtered_df = filtered_df[filtered_df['Product Type'] == selected_product]
        
    if selected_brand != 'All':
        filtered_df = filtered_df[filtered_df['Brand: Name'] == selected_brand]
    
    return filtered_df

def calculate_statistics(df):
    """Calculate statistics for each dealer category"""
    if df.empty:
        st.warning("No data available with the selected filters")
        return None
    
    # Group by dealer category and calculate statistics
    stats = []
    categories = sorted(df['Account: Dealer Category'].unique())
    
    for category in categories:
        category_df = df[df['Account: Dealer Category'] == category]
        prices = category_df['Whole Sale Price'].dropna()
        
        if not prices.empty:
            try:
                modal_value = mode(prices)
            except:
                modal_value = np.nan  # If no unique mode exists
            
            stat = {
                'Dealer Category': category,
                'Count': len(prices),
                'Minimum': prices.min(),
                'Maximum': prices.max(),
                'Average': prices.mean(),
                'Median': median(prices),
                'Mode': modal_value
            }
        else:
            stat = {
                'Dealer Category': category,
                'Count': 0,
                'Minimum': np.nan,
                'Maximum': np.nan,
                'Average': np.nan,
                'Median': np.nan,
                'Mode': np.nan
            }
        
        stats.append(stat)
    
    stats_df = pd.DataFrame(stats)
    return stats_df

def display_interactive_table(stats_df, filtered_df):
    """Display an interactive table with statistics and drill-down capability"""
    if stats_df is None or stats_df.empty:
        return
    
    st.subheader("Wholesale Price Statistics by Dealer Category")
    
    # Format the statistics dataframe for display
    display_df = stats_df.copy()
    display_df['Average'] = display_df['Average'].round(2)
    
    # Create tabs for the main table and detailed view
    tab1, tab2 = st.tabs(["Summary Statistics", "Detailed View"])
    
    with tab1:
        st.dataframe(display_df.set_index('Dealer Category'), use_container_width=True)
        
        # Create a bar chart comparing average prices
        fig = px.bar(
            display_df, 
            x='Dealer Category', 
            y='Average',
            title='Average Wholesale Price by Dealer Category',
            labels={'Average': 'Average Wholesale Price'},
            text_auto='.2f'
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        # Create dealer category selector for detailed view
        selected_category = st.selectbox(
            "Select a dealer category to see details",
            options=stats_df['Dealer Category'].tolist()
        )
        
        # Show detailed information for the selected category
        category_df = filtered_df[filtered_df['Account: Dealer Category'] == selected_category]
        
        if not category_df.empty:
            st.subheader(f"Detailed Entries for '{selected_category}' Category")
            
            # Group by officer and show their entries
            officers = category_df['Owner: Full Name'].unique()
            
            officer_data = []
            for officer in officers:
                officer_df = category_df[category_df['Owner: Full Name'] == officer]
                
                for _, row in officer_df.iterrows():
                    officer_data.append({
                        'Officer Name': officer,
                        'District': row['District: Name'],
                        'Checkin Date': row['checkin date'].date(),
                        'Product Type': row['Product Type'],
                        'Brand': row['Brand: Name'],
                        'Wholesale Price': row['Whole Sale Price']
                    })
            
            officer_df = pd.DataFrame(officer_data)
            st.dataframe(officer_df, use_container_width=True)
            
            # Create visualization of officer price inputs
            if len(officer_data) > 0:
                fig = px.box(
                    officer_df, 
                    x='Officer Name', 
                    y='Wholesale Price',
                    title=f'Wholesale Price Distribution by Officer for {selected_category} Category',
                    points="all"
                )
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No data available for '{selected_category}' category with the current filters")

def main():
    """Main function to run the dashboard"""
    st.title("Wholesale Price Analysis Dashboard")
    st.write("Upload your dataset to analyze dealer wholesale prices")
    
    # Load data
    df = load_data()
    
    if df is not None:
        # Display dataset overview
        st.subheader("Dataset Overview")
        st.write(f"Total records: {len(df)}")
        st.write(f"Date range: {df['checkin date'].min().date()} to {df['checkin date'].max().date()}")
        st.write(f"Number of districts: {df['District: Name'].nunique()}")
        st.write(f"Number of dealer categories: {df['Account: Dealer Category'].nunique()}")
        
        # Create and apply filters
        filtered_df = create_filters(df)
        
        # Calculate statistics
        stats_df = calculate_statistics(filtered_df)
        
        # Display interactive table
        display_interactive_table(stats_df, filtered_df)
        
        # Option to download filtered data
        if not filtered_df.empty:
            st.subheader("Download Filtered Data")
            csv = filtered_df.to_csv(index=False)
            st.download_button(
                label="Download as CSV",
                data=csv,
                file_name="filtered_price_data.csv",
                mime="text/csv"
            )

if __name__ == "__main__":
    main()
