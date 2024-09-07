import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from scipy import stats
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
from openpyxl import load_workbook
from PIL import Image

# Initialize session state
if 'df' not in st.session_state:
    st.session_state.df = None
if 'week_names_input' not in st.session_state:
    st.session_state.week_names_input = []
if 'desired_diff_input' not in st.session_state:
    st.session_state.desired_diff_input = {}

def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]  # Use week name from user input
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df,
                                  week_data,
                                  left_index=True,
                                  right_index=True)
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands, desired_diff, week_names, download_pdf=False, diff_week=1):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    if download_pdf:
        pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf")
    for i, district_name in enumerate(district_names):
        plt.figure(figsize=(10, 8))
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        for brand in brands:
            brand_prices = []
            for week_name in week_names:  # Use week_names directly
                column_name = f"{brand} ({week_name})"
                if column_name in district_df.columns:
                    price = district_df[column_name].iloc[0]
                    brand_prices.append(price)
                else:
                    brand_prices.append(np.nan)
            valid_prices = [p for p in brand_prices if not np.isnan(p)]
            if len(valid_prices) > diff_week:
                price_diff = valid_prices[-1] - valid_prices[diff_week]
            else:
                price_diff = np.nan
            price_diffs.append(price_diff)
            line, = plt.plot(week_names,  # Use week_names directly
                             brand_prices,
                             marker='o',
                             linestyle='-',
                             label=f"{brand} ({price_diff:.0f})")
            for week, price in zip(week_names, brand_prices):  # Use week_names directly
                if not np.isnan(price):
                    plt.text(week, price, str(round(price)), fontsize=10)
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        plt.ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        # Add region name above the title only for the first district
        if i == 0:
            #plt.title(f"{region_name}\n{district_name} - Brands Price Trend", weight='bold',fontsize=16)
            plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)  # Added region name using plt.text
            plt.title(f"{district_name} - Brands Price Trend", weight='bold') # Keep the original title without region name
        else:
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()

        text_str = ''
        if benchmark_brands:
           brand_texts = []
           max_left_length = 0  # Store text for each brand separately
           for benchmark_brand in benchmark_brands:
               jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
               benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
               actual_diff = np.nan  # Initialize actual_diff with NaN
        if jklc_prices and benchmark_prices:
            for i in range(len(jklc_prices) - 1, -1, -1):
                if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                    actual_diff = jklc_prices[i] - benchmark_prices[i]
                    break
        desired_diff_str = f" ({desired_diff[benchmark_brand]:.0f} Rs.)" if benchmark_brand in desired_diff and desired_diff[benchmark_brand] is not None else ""
        brand_text = [f"Benchmark Brand: {benchmark_brand}{desired_diff_str}", f"Actual Diff: {actual_diff:+.2f} Rs."]
        brand_texts.append(brand_text)
        max_left_length = max(max_left_length, len(brand_text[0]))
    num_brands = len(brand_texts)
    if num_brands == 1:
        text_str = "\n".join(brand_texts[0])
    elif num_brands > 1:
        half_num_brands = num_brands // 2
        left_side = brand_texts[:half_num_brands]
        right_side = brand_texts[half_num_brands:]
        lines = []
        for i in range(2):
            left_text = left_side[0][i] if i < len(left_side[0]) else ""
            right_text = right_side[0][i] if i < len(right_side[0]) else ""
            lines.append(f"{left_text.ljust(max_left_length)} \u2502 {right_text.rjust(max_left_length)}")
        text_str = "\n".join(lines)
    plt.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=plt.gca().transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
    plt.subplots_adjust(bottom=0.25)
    if download_pdf:
        pdf.savefig()
    buf = BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    b64_data = base64.b64encode(buf.getvalue()).decode()
    st.markdown(f'<a download="district_plot_{district_name}.png" href="data:image/png;base64,{b64_data}">Download Plot as PNG</a>', unsafe_allow_html=True)
    plt.show()    
if download_pdf:
    pdf.close()
    with open("district_plots.pdf", "rb") as f:
        pdf_data = f.read()
    b64_pdf = base64.b64encode(pdf_data).decode()
    st.markdown(f'<a download="{region_name}.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>', unsafe_allow_html=True)   

def main():
    uploaded_file = st.file_uploader("Choose a file")
    if uploaded_file is not None:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Get the week names from the user
        num_weeks = len([col for col in df.columns if 'JKLC' in col]) // 6
        week_names_input = []
        for i in range(num_weeks):
            week_name = st.text_input(f"Week {i+1}:")
            week_names_input.append(week_name)
        
        # Transform the data
        df = transform_data(df, week_names_input)
        
        # Get the zone names
        zone_names = df["Zone"].unique().tolist()
        zone_name = st.selectbox("Select Zone:", zone_names)
        
        # Filter the data by zone
        filtered_df = df[df["Zone"] == zone_name]
        
        # Get the region names
        region_names = filtered_df["REGION"].unique().tolist()
        region_name = st.selectbox("Select Region:", region_names)
        
        # Filter the data by region
        filtered_df = filtered_df[filtered_df["REGION"] == region_name]
        
        # Get the district names
        district_names = filtered_df["Dist Name"].unique().tolist()
        district_name = st.multiselect("Select District:", district_names)
        
        # Get the benchmark brands
       district_names = filtered_df["Dist Name"].unique().tolist()
       district_name = st.multiselect("Select District:", district_names)

# Get the benchmark brands
       benchmark_brands = ['UTCL', 'JKS', 'Ambuja', 'Wonder', 'Shree']
       benchmark_brand = st.multiselect("Select Benchmark Brands:", benchmark_brands)

# Get the desired difference for each benchmark brand
       desired_diff = {}
       for brand in benchmark_brand:
           desired_diff[brand] = st.number_input(f"Desired Difference for {brand}:")
    
# Plot the district grap
plot_district_graph(df, district_name, benchmark_brand, desired_diff, week_names_input)
desired_diff = {}
for brand in benchmark_brand:
    desired_diff[brand] = st.number_input(f"Desired Difference for {brand}:")

# Plot the district graph
if st.button("Plot District Graph"):
    plot_district_graph(df, district_name, benchmark_brand, desired_diff, week_names_input)
    if st.checkbox("Download PDF"):
        plot_district_graph(df, district_name, benchmark_brand, desired_diff, week_names_input, download_pdf=True)
if __name__ == "__main__":
    main()
