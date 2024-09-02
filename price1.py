
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import base64
from io import BytesIO

df = None
desired_diff_input = {}  # Initialize desired_diff_input as a dictionary

def transform_data(df):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    month_names = ['June', 'July', 'August', 'September', 'October', 'November',
                   'December', 'January', 'February', 'March', 'April', 'May']
    month_index = 0
    week_counter = 1
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        if i == 0:  # Special handling for June
            week_name = month_names[month_index]
            month_index += 1
        elif i == 1:
            week_name = month_names[month_index]
            month_index += 1
        else:
            week_name = f"W-{week_counter} {month_names[month_index]}"
            if week_counter == 4:
                week_counter = 1
                month_index += 1
            else:
                week_counter += 1
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True) # Replace 0 with NaN in week_data
        transformed_df = pd.merge(transformed_df,
                                  week_data,
                                  left_index=True,
                                  right_index=True)
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands, desired_diff):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    week_names = list(set([col.split(' (')[1].split(')')[0] for col in df.columns if '(' in col]))
    def sort_week_names(week_name):
        if ' ' in week_name:  # Check for week names with month
            week, month = week_name.split()
            week_num = int(week.split('-')[1])
        else:  # Handle June without week number
            week_num = 0
            month = week_name
        month_order = ['June', 'July', 'August', 'September', 'October', 'November',
            'December', 'January', 'February', 'March', 'April', 'May']
        month_num = month_order.index(month)
        return month_num * 10 + week_num
    week_names.sort(key=sort_week_names)  # Sort using the custom function
    for district_name in district_names:  # Iterate over selected district names
        fig, ax = plt.subplots(figsize=(10, 6)) # Increased figure height to accommodate the text box
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        stats_table_data = {}
        for brand in brands:
           brand_prices = []
           for week_name in week_names:
                column_name = f"{brand} ({week_name})"
                if column_name in district_df.columns:
                    price = district_df[column_name].iloc[0]
                    brand_prices.append(price)
                else:
                    brand_prices.append(np.nan)
           valid_prices = [p for p in brand_prices if not np.isnan(p)]
           if len(valid_prices) >= 2:
                price_diff = valid_prices[-1] - valid_prices[1]
           else:
                price_diff = np.nan
           price_diffs.append(price_diff)
           line, = ax.plot(week_names,
                             brand_prices,
                             marker='o',
                             linestyle='-',
                             label=f"{brand} ({price_diff:.0f})")
           for week, price in zip(week_names, brand_prices):
                if not np.isnan(price):
                    ax.text(week, price, str(round(price)), fontsize=10)
           if valid_prices:
                stats_table_data[brand] = {
                    'Min': np.min(valid_prices),
                    'Max': np.max(valid_prices),
                    'Average': np.mean(valid_prices),
                    'Median': np.median(valid_prices),
                    'First Quartile': np.percentile(valid_prices, 25),
                    'Third Quartile': np.percentile(valid_prices, 75),
                    'Variance': np.var(valid_prices),
                    'Skewness': pd.Series(valid_prices).skew(),
                    'Kurtosis': pd.Series(valid_prices).kurtosis()
                }
           else:
                stats_table_data[brand] = {
                    'Min': np.nan,
                    'Max': np.nan,
                    'Average': np.nan,
                    'Median': np.nan,
                    'First Quartile': np.nan,
                    'Third Quartile': np.nan,
                    'Variance': np.nan,
                    'Skewness': np.nan,
                    'Kurtosis': np.nan
                }
        ax.grid(False)
        ax.set_xlabel('Month/Week',weight='bold')
        ax.set_ylabel('Whole Sale Price(in Rs.)',weight='bold')
        ax.set_title(f"{district_name} - Brands Price Trend",weight='bold')
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6,prop={'weight': 'bold'})
        fig.tight_layout()
        stats_table = pd.DataFrame(stats_table_data).transpose().round(2)
        st.write(stats_table)
        # Create a text box below the graph
        text_str = ''
        if benchmark_brands:
            brand_texts = []  # Store text for each brand separately
            max_left_length = 0 # Initialize to keep track of the longest left side text
            for benchmark_brand in benchmark_brands:
                jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
                benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
                actual_diff = np.nan  # Initialize actual_diff with NaN
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices)-1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break

                desired_diff_str = f" ({desired_diff[benchmark_brand].value:+.2f} Rs.)" if benchmark_brand in desired_diff and desired_diff[benchmark_brand].value is not None else ""
                brand_text = [f"Benchmark Brand: {benchmark_brand}{desired_diff_str}", f"Actual Diff: {actual_diff:+.2f} Rs."]
                brand_texts.append(brand_text)
                max_left_length = max(max_left_length, len(brand_text[0])) # Update max_left_length if current brand text is longer

            # Join brand texts with a vertical line separator
            num_brands = len(brand_texts)
            if num_brands ==1:
                text_str = "\n".join(brand_texts[0])
            elif num_brands > 1:
                half_num_brands = num_brands // 2
                left_side = brand_texts[:half_num_brands]
                right_side = brand_texts[half_num_brands:]

                lines = []
                for i in range(2): # Iterate over the 2 lines of each brand
                    left_text = left_side[0][i] if i < len(left_side[0]) else "" 
                    right_text = right_side[0][i] if i < len(right_side[0]) else ""
                    lines.append(f"{left_text.ljust(max_left_length)} \u2502 {right_text.rjust(max_left_length)}") # Pad with spaces
                text_str = "\n".join(lines)

        ax.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=ax.transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
        fig.subplots_adjust(bottom=0.2)

        buf = BytesIO()
        fig.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        b64_data = base64.b64encode(buf.getvalue()).decode()
        href = f'<a href="data:image/png;base64,{b64_data}" download="district_plot_{district_name}.png">Download Plot as PNG</a>'
        st.markdown(href, unsafe_allow_html=True)
        st.pyplot(fig)

def on_button_click(uploaded_file):
    if uploaded_file:
        try:
            file_name = uploaded_file.name
            global df
            df = pd.read_excel(uploaded_file, skiprows=2)
            df = transform_data(df)
            zone_names = df["Zone"].unique().tolist()
            zone_dropdown.options = zone_names
            st.success(f"Uploaded file: {file_name}")  # Print the file name
            create_interactive_plot(df)
        except Exception as e:
            st.error(f"Error reading file: {e}")

def on_zone_change(selected_zone):
    global df
    filtered_df = df[df["Zone"] == selected_zone]
    region_names = filtered_df["REGION"].unique().tolist()
    region_dropdown.options = region_names
    district_dropdown.options = []  # Clear district options

def on_region_change(selected_region):
    global df
    filtered_df = df[df["REGION"] == selected_region]
    district_names = filtered_df["Dist Name"].unique().tolist()
    district_dropdown.options = district_names

def on_district_change(selected_districts):
    global df, desired_diff_input
    if selected_districts:  # Only update if districts are selected
        all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
        benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
        benchmark_dropdown.options = benchmark_brands

def create_interactive_plot(df):
    global desired_diff_input
    all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
    desired_diff_input = {brand: st.number_input(f'Desired Diff for {brand}:', value=0) for brand in benchmark_brands}

    if st.button("Generate Plot"):
        selected_districts = district_dropdown.value
        selected_benchmark_brands = benchmark_dropdown.value
        if selected_districts and selected_benchmark_brands:
            plot_district_graph(df, selected_districts, selected_benchmark_brands, desired_diff_input)

# Streamlit app layout
st.title("District-wise Brand Price Trend Analysis")

uploaded_file = st.file_uploader("Upload Excel", type=['xlsx'])
if uploaded_file is not None:
    on_button_click(uploaded_file)

zone_dropdown = st.selectbox("Select Zone", [])
region_dropdown = st.selectbox("Select Region", [])
district_dropdown = st.multiselect("Select District", [])
benchmark_dropdown = st.multiselect("Select Benchmark Brands", [])

if zone_dropdown:
    on_zone_change(zone_dropdown)
if region_dropdown:
    on_region_change(region_dropdown)
if district_dropdown:
    on_district_change(district_dropdown)

create_interactive_plot(df)
