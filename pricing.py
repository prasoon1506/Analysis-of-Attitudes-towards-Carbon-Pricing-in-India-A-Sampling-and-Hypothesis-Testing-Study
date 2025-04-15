import pandas as pd
import numpy as np
import streamlit as st
import plotly.express as px
from datetime import datetime, timedelta
import plotly.graph_objects as go
from statistics import mode, median
import io
st.set_page_config(layout="wide", page_title="Dealer Price Analysis Dashboard")
def load_data():
    st.sidebar.title("Data Upload")
    uploaded_file = st.sidebar.file_uploader("Upload your dataset (CSV or Excel)", type=["csv", "xlsx", "xls"])
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file)
            else:
                df = pd.read_excel(uploaded_file)
            required_columns = ['District: Name', 'checkin date', 'Product Type', 'Brand: Name','Account: Dealer Category', 'Whole Sale Price', 'Owner: Full Name']
            missing_columns = [col for col in required_columns if col not in df.columns]
            if missing_columns:
                st.error(f"Missing required columns: {', '.join(missing_columns)}")
                return None
            try:
                df['checkin date'] = pd.to_datetime(df['checkin date'])
            except Exception as e:
                st.error(f"Error converting 'checkin date' to datetime: {e}")
                return None
            df['Account: Dealer Category'] = df['Account: Dealer Category'].fillna('NaN')
            df['Whole Sale Price'] = pd.to_numeric(df['Whole Sale Price'], errors='coerce')
            return df
        except Exception as e:
            st.error(f"Error loading data: {e}")
            return None
    return None
def create_filters(df):
    st.sidebar.header("Filters")
    df = df.copy()
    district_values = df['District: Name'].fillna('Unknown').unique().tolist()
    district_options = ['All'] + sorted(district_values)
    selected_district = st.sidebar.selectbox("Select District", district_options)
    min_date = df['checkin date'].min().date()
    max_date = df['checkin date'].max().date()
    date_selection_type = st.sidebar.radio("Date Selection", ["Single Date", "Date Range"])
    if date_selection_type == "Single Date":
        selected_date = st.sidebar.date_input("Select Checkin Date", min_date,min_value=min_date, max_value=max_date)
        date_filter = (df['checkin date'].dt.date == selected_date)
    else:
        default_end_date = min(max_date, min_date + timedelta(days=3))
        date_range = st.sidebar.date_input("Select Date Range",[min_date, default_end_date],min_value=min_date,max_value=max_date)
        if len(date_range) == 2:
            start_date, end_date = date_range
            date_filter = (df['checkin date'].dt.date >= start_date) & (df['checkin date'].dt.date <= end_date)
        else:
            st.sidebar.warning("Please select both start and end dates")
            date_filter = pd.Series(True, index=df.index)
    product_values = df['Product Type'].fillna('Unknown').unique().tolist()
    product_options = ['All'] + sorted(product_values)
    selected_product = st.sidebar.selectbox("Select Product Type", product_options)
    brand_values = df['Brand: Name'].fillna('Unknown').unique().tolist()
    brand_options = ['All'] + sorted(brand_values)
    selected_brand = st.sidebar.selectbox("Select Brand", brand_options)
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
    if df.empty:
        st.warning("No data available with the selected filters")
        return None
        
    stats = []
    categories = sorted(df['Account: Dealer Category'].unique())
    
    for category in categories:
        category_df = df[df['Account: Dealer Category'] == category]
        prices = category_df['Whole Sale Price'].dropna()
        
        if not prices.empty:
            try:
                modal_value = mode(prices)
            except:
                modal_value = np.nan
                
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
    if stats_df is None or stats_df.empty:
        return
        
    st.subheader("Wholesale Price Statistics by Dealer Category")
    
    display_df = stats_df.copy()
    display_df['Average'] = display_df['Average'].round(2)
    
    tab1, tab2 = st.tabs(["Summary Statistics", "Detailed View"])
    
    with tab1:
        st.dataframe(display_df.set_index('Dealer Category'), use_container_width=True)
        
        fig = px.bar(display_df,
                    x='Dealer Category',
                    y='Average',
                    title='Average Wholesale Price by Dealer Category',
                    labels={'Average': 'Average Wholesale Price'},
                    text_auto='.2f')
        
        st.plotly_chart(fig, use_container_width=True)
    
    with tab2:
        selected_category = st.selectbox("Select a dealer category to see details",
                                       options=stats_df['Dealer Category'].tolist())
        
        category_df = filtered_df[filtered_df['Account: Dealer Category'] == selected_category]
        
        if not category_df.empty:
            st.subheader(f"Detailed Entries for '{selected_category}' Category")
            
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
            
            if len(officer_data) > 0:
                fig = px.box(officer_df,
                           x='Officer Name',
                           y='Wholesale Price',
                           title=f'Wholesale Price Distribution by Officer for {selected_category} Category',
                           points="all")
                
                st.plotly_chart(fig, use_container_width=True)
        else:
            st.info(f"No data available for '{selected_category}' category with the current filters")
def generate_excel_report(df):
    st.subheader("Generate Professional Excel Report")
    
    all_districts = sorted(df['District: Name'].dropna().unique().tolist())
    selected_districts = st.multiselect("Select Districts to Include in Report", ["All"] + all_districts)
    
    if not selected_districts:
        st.info("Please select at least one district.")
        return
        
    if "All" in selected_districts:
        selected_districts = all_districts
    
    all_brands = sorted(df['Brand: Name'].dropna().unique().tolist())
    selected_brands = st.multiselect("Select Brands to Include in Report", all_brands)
    
    if not selected_brands:
        st.info("Please select at least one brand.")
        return
    
    df = df.copy()
    df['Brand: Name'] = df['Brand: Name'].replace({'Shree Cement': 'Shree'})
    df['Account: Dealer Category'] = df['Account: Dealer Category'].fillna('NaN')
    df['checkin date'] = pd.to_datetime(df['checkin date'])
    
    filtered_df = df[df['District: Name'].isin(selected_districts) & df['Brand: Name'].isin(selected_brands)].copy()
    
    date_range = pd.date_range(filtered_df['checkin date'].min(), filtered_df['checkin date'].max())
    date_columns = [d.date() for d in date_range]
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for brand in selected_brands:
            brand_df = filtered_df[filtered_df['Brand: Name'] == brand]
            rows = []
            
            for district in selected_districts:
                district_df = brand_df[brand_df['District: Name'] == district]
                officers = sorted(district_df['Owner: Full Name'].dropna().unique().tolist())
                
                for officer in officers:
                    officer_df = district_df[district_df['Owner: Full Name'] == officer]
                    categories = sorted(officer_df['Account: Dealer Category'].unique().tolist())
                    
                    # Process each dealer category for this officer
                    for category in categories:
                        row = {'District': district, 'Officer': officer, 'Dealer Category': category}
                        cat_df = officer_df[officer_df['Account: Dealer Category'] == category]
                        
                        for d in date_columns:
                            day_data = cat_df[cat_df['checkin date'].dt.date == d]['Whole Sale Price']
                            if len(day_data) == 1:
                                row[d.strftime("%d-%b")] = day_data.iloc[0]
                            elif len(day_data) == 2:
                                if day_data.iloc[0] != day_data.iloc[1]:
                                    row[d.strftime("%d-%b")] = ', '.join(map(str, day_data))
                                else:
                                    row[d.strftime("%d-%b")] = day_data.iloc[0]
                            elif len(day_data) > 2:
                                try:
                                    row[d.strftime("%d-%b")] = mode(day_data)
                                except:
                                    row[d.strftime("%d-%b")] = np.nan
                            else:
                                row[d.strftime("%d-%b")] = np.nan
                                
                        sorted_prices = cat_df.sort_values('checkin date')['Whole Sale Price'].dropna()
                        if len(sorted_prices) >= 2:
                            row['Change'] = sorted_prices.iloc[-1] - sorted_prices.iloc[0] \
                                if sorted_prices.iloc[-1] != sorted_prices.iloc[0] else 0
                        elif len(sorted_prices) == 1:
                            row['Change'] = '-'  # Not enough data
                        else:
                            row['Change'] = np.nan
                            
                        row['Total Inputs'] = len(cat_df)
                        rows.append(row)
                    
                    # Only add Overall line if officer has more than one dealer category
                    if len(categories) > 1:
                        row = {'District': district, 'Officer': officer, 'Dealer Category': 'Overall'}
                        
                        for d in date_columns:
                            day_data = officer_df[officer_df['checkin date'].dt.date == d]['Whole Sale Price']
                            if len(day_data) == 1:
                                row[d.strftime("%d-%b")] = day_data.iloc[0]
                            elif len(day_data) == 2:
                                if day_data.iloc[0] != day_data.iloc[1]:
                                    row[d.strftime("%d-%b")] = ', '.join(map(str, day_data))
                                else:
                                    row[d.strftime("%d-%b")] = day_data.iloc[0]
                            elif len(day_data) > 2:
                                try:
                                    row[d.strftime("%d-%b")] = mode(day_data)
                                except:
                                    row[d.strftime("%d-%b")] = np.nan
                            else:
                                row[d.strftime("%d-%b")] = np.nan
                                
                        full_data = officer_df.sort_values('checkin date')['Whole Sale Price'].dropna()
                        if len(full_data) >= 2:
                            row['Change'] = full_data.iloc[-1] - full_data.iloc[0] \
                                if full_data.iloc[-1] != full_data.iloc[0] else 0
                        elif len(full_data) == 1:
                            row['Change'] = '-'
                        else:
                            row['Change'] = np.nan
                            
                        row['Total Inputs'] = len(officer_df)
                        rows.append(row)
                
                # Process each dealer category for OVERALL (across all officers)
                categories = sorted(district_df['Account: Dealer Category'].unique().tolist())
                for category in categories:
                    row = {'District': district, 'Officer': 'OVERALL', 'Dealer Category': category}
                    cat_df = district_df[district_df['Account: Dealer Category'] == category]
                    
                    for d in date_columns:
                        day_data = cat_df[cat_df['checkin date'].dt.date == d]['Whole Sale Price']
                        if len(day_data) == 1:
                            row[d.strftime("%d-%b")] = day_data.iloc[0]
                        elif len(day_data) == 2:
                            if day_data.iloc[0] != day_data.iloc[1]:
                                row[d.strftime("%d-%b")] = ', '.join(map(str, day_data))
                            else:
                                row[d.strftime("%d-%b")] = day_data.iloc[0]
                        elif len(day_data) > 2:
                            try:
                                row[d.strftime("%d-%b")] = mode(day_data)
                            except:
                                row[d.strftime("%d-%b")] = np.nan
                        else:
                            row[d.strftime("%d-%b")] = np.nan
                            
                    sorted_prices = cat_df.sort_values('checkin date')['Whole Sale Price'].dropna()
                    if len(sorted_prices) >= 2:
                        row['Change'] = sorted_prices.iloc[-1] - sorted_prices.iloc[0] \
                            if sorted_prices.iloc[-1] != sorted_prices.iloc[0] else 0
                    elif len(sorted_prices) == 1:
                        row['Change'] = '-'
                    else:
                        row['Change'] = np.nan
                        
                    row['Total Inputs'] = len(cat_df)
                    rows.append(row)
                
                # Overall line for the entire district
                row = {'District': district, 'Officer': 'OVERALL', 'Dealer Category': 'Overall'}
                
                for d in date_columns:
                    day_data = district_df[district_df['checkin date'].dt.date == d]['Whole Sale Price']
                    if len(day_data) == 1:
                        row[d.strftime("%d-%b")] = day_data.iloc[0]
                    elif len(day_data) == 2:
                        if day_data.iloc[0] != day_data.iloc[1]:
                            row[d.strftime("%d-%b")] = ', '.join(map(str, day_data))
                        else:
                            row[d.strftime("%d-%b")] = day_data.iloc[0]
                    elif len(day_data) > 2:
                        try:
                            row[d.strftime("%d-%b")] = mode(day_data)
                        except:
                            row[d.strftime("%d-%b")] = np.nan
                    else:
                        row[d.strftime("%d-%b")] = np.nan
                        
                full_data = district_df.sort_values('checkin date')['Whole Sale Price'].dropna()
                if len(full_data) >= 2:
                    row['Change'] = full_data.iloc[-1] - full_data.iloc[0] \
                        if full_data.iloc[-1] != full_data.iloc[0] else 0
                elif len(full_data) == 1:
                    row['Change'] = '-'
                else:
                    row['Change'] = np.nan
                    
                row['Total Inputs'] = len(district_df)
                rows.append(row)
            
            brand_report_df = pd.DataFrame(rows)
            brand_report_df.to_excel(writer, sheet_name=brand, index=False)
            
            workbook = writer.book
            worksheet = writer.sheets[brand]
            
            # Define formats
            header_format = workbook.add_format({'bold': True, 'bg_color': '#D9EAD3', 'border': 1, 'align': 'center'})
            cell_format = workbook.add_format({'border': 1, 'align': 'center'})
            number_format = workbook.add_format({'border': 1, 'align': 'center', 'num_format': '0'})
            change_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#F4CCCC', 'bold': True})
            total_input_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#FFEB9C'})
            overall_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#BDD7EE', 'bold': True})
            officer_overall_format = workbook.add_format({'border': 1, 'align': 'center', 'bg_color': '#DCE6F1', 'bold': True})
            
            # Write headers
            for col_num, value in enumerate(brand_report_df.columns):
                worksheet.write(0, col_num, value, header_format)
            
            # Write data with appropriate formatting
            for row_num, row in enumerate(brand_report_df.values):
                is_overall_row = row[1] == 'OVERALL'
                is_officer_overall = row[2] == 'Overall' and row[1] != 'OVERALL'
                
                for col_num, value in enumerate(row):
                    col_name = brand_report_df.columns[col_num]
                    
                    if is_overall_row:
                        base_format = overall_format
                    elif is_officer_overall:
                        base_format = officer_overall_format
                    elif col_name == 'Change' and isinstance(value, (int, float)) and not pd.isna(value) and np.isfinite(value):
                        base_format = change_format
                    elif col_name == 'Total Inputs' and isinstance(value, (int, float)) and not pd.isna(value) and np.isfinite(value):
                        base_format = total_input_format
                    elif isinstance(value, (int, float)) and not pd.isna(value) and np.isfinite(value):
                        base_format = number_format
                    else:
                        base_format = cell_format
                    
                    if pd.isna(value):
                        worksheet.write(row_num + 1, col_num, '', base_format)
                    else:
                        worksheet.write(row_num + 1, col_num, value, base_format)
            
            # Auto-fit column widths
            for col_num, col_name in enumerate(brand_report_df.columns):
                # Find the maximum width needed
                max_len = len(str(col_name))
                for row_num, row in enumerate(brand_report_df.values):
                    if not pd.isna(row[col_num]):
                        cell_len = len(str(row[col_num]))
                        if cell_len > max_len:
                            max_len = cell_len
                # Add a buffer for better appearance
                adjusted_width = max_len + 2
                # Set the column width (max 50 chars to prevent excessive width)
                worksheet.set_column(col_num, col_num, min(adjusted_width, 50))
            
            # Add filters to the header row
            worksheet.autofilter(0, 0, len(brand_report_df), len(brand_report_df.columns) - 1)
            
            # Districts to be pre-selected in filter
            focus_districts = [
                "Z0605_Ahmadabad", "Z0616_Surat", "Z0703_Gurugram", "Z1226_Indore", 
                "Z1230_Balaghat", "Z1329_Nagpur", "Z1804_Sambalpur", "Z1810_Khorda", 
                "Z1909_Bathinda", "Z2013_Udaipur", "Z2020_Jaipur", "Z2405_Ghaziabad", 
                "Z2470_Varanasi", "Z3001_East", "Z3505_Dehradun", "Z3506_Haridwar", 
                "Z3406_Palamu", "Z3302_Raipur"
            ]
            
            # Get districts present in this sheet
            present_districts = brand_report_df['District'].unique().tolist()
            
            # Find which focused districts are present in this sheet
            present_focus_districts = [d for d in focus_districts if d in present_districts]
            
            # If we have focus districts in this sheet, set up the filter
            if present_focus_districts and len(present_focus_districts) < len(present_districts):
                try:
                    # Create a list of only the focus districts to show
                    filter_items = []
                    for district in present_focus_districts:
                        filter_items.append(district)
                    
                    # Apply the filter to show only selected districts
                    worksheet.filter_column_list(0, filter_items)
                except Exception as e:
                    # If filtering fails, just log the error and continue without filter
                    st.warning(f"Could not apply district filter: {str(e)}")
            
            worksheet.freeze_panes(1, 0)
    
    st.success("Excel report is ready.")
    st.download_button(
        "Download Report as Excel",
        data=output.getvalue(),
        file_name="dealer_price_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

def main():
    st.title("Dealer Price Analysis Dashboard")
    st.write("Upload your dataset to analyze dealer wholesale prices")
    
    df = load_data()
    
    if df is not None:
        st.subheader("Dataset Overview")
        st.write(f"Total records: {len(df)}")
        st.write(f"Date range: {df['checkin date'].min().date()} to {df['checkin date'].max().date()}")
        st.write(f"Number of districts: {df['District: Name'].nunique()}")
        st.write(f"Number of dealer categories: {df['Account: Dealer Category'].nunique()}")
        
        filtered_df = create_filters(df)
        stats_df = calculate_statistics(filtered_df)
        display_interactive_table(stats_df, filtered_df)
        
        generate_excel_report(df)
        
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
