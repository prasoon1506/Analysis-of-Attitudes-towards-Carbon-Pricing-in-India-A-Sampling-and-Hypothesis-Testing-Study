pip install -r requirements.txt
import openpyxl
import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from tqdm import tqdm
import xgboost as xgb
from scipy import stats

# Set page title and favicon
st.set_page_config(page_title="Brand Price Trend Analysis", page_icon=":chart_with_upwards_trend:")

# --- Data Transformation Function ---
def transform_data(df):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    month_names = ['June', 'July', 'August', 'September', 'October', 'November', 'December',
                   'January', 'February', 'March', 'April', 'May']
    month_index = 0
    week_counter = 1

    for i in tqdm(range(num_weeks), desc="Transforming data"):
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
            col: f"{brand} ({week_name})" for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True)
    return transformed_df

# --- District Graph Plotting Function ---
def plot_district_graph(df, district_names, benchmark_brands, desired_diff):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    week_names = list(set([col.split(' (')[1].split(')')[0] for col in df.columns if '(' in col]))

    def sort_week_names(week_name):
        if ' ' in week_name:
            week, month = week_name.split()
            week_num = int(week.split('-')[1])
        else:
            week_num = 0
            month = week_name
        month_order = ['June', 'July', 'August', 'September', 'October', 'November',
                       'December', 'January', 'February', 'March', 'April', 'May']
        month_num = month_order.index(month)
        return month_num * 10 + week_num

    week_names.sort(key=sort_week_names)

    all_stats_tables = []  # List to store stats tables for all districts
    all_predictions = []  # List to store predictions for all districts

    for district_name in tqdm(district_names, desc="Processing districts"):
        fig, ax = plt.subplots(figsize=(10, 8))
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        stats_table_data = {}
        predictions = {}
    for brand in brands:
            brand_prices = []
            for week_name in week_names:
                column_name = f"{brand} ({week_name})"
                if column_name in district_df.columns:
                    price = district_df[column_name].iloc[0]
                    brand_prices.append(price)
                else:
                    brand_prices.append(np.nan)

            # Price Trend Plotting and Stats Calculation 
            valid_prices = [p for p in brand_prices if not np.isnan(p)]
            if valid_prices:
                price_diff = valid_prices[-1] - valid_prices[0]
            else:
                price_diff = np.nan
            price_diffs.append(price_diff)
            ax.plot(week_names, brand_prices, marker='o', linestyle='-', label=f"{brand} ({price_diff:.0f})")

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
             # ... (previous code for stats_table_data)

            # Prediction using XGBoost
            if len(valid_prices) > 2:  # Need at least 3 data points for prediction
                train_data = np.array(range(len(valid_prices))).reshape(-1, 1)
                train_labels = np.array(valid_prices)
                model = xgb.XGBRegressor(objective='reg:squarederror')
                model.fit(train_data, train_labels)
                next_week = len(valid_prices)
                prediction = model.predict(np.array([[next_week]]))

                # Calculate confidence interval
                errors = abs(model.predict(train_data) - train_labels)
                confidence = 0.95
                n = len(valid_prices)
                t_crit = stats.t.ppf((1 + confidence) / 2, n - 1)
                margin_of_error = t_crit * errors.std() / np.sqrt(n)
                confidence_interval = (prediction - margin_of_error,
                                       prediction + margin_of_error)

                predictions[brand] = {
                    'Prediction': prediction[0],
                    'Confidence Interval': confidence_interval
                }
            else:
                predictions[brand] = {
                    'Prediction': np.nan,
                    'Confidence Interval': (np.nan, np.nan)
                }

        # Plotting and Text Box 
    ax.grid(False)
    ax.set_xlabel('Month/Week', weight='bold')
    ax.set_ylabel('Whole Sale Price(in Rs.)', weight='bold')
    ax.set_title(f"{district_name} - Brands Price Trend", weight='bold')
    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})

        # Create stats table and predictions table
    stats_table = pd.DataFrame(stats_table_data).transpose().round(2)
    st.write(stats_table)
    all_stats_tables.append(stats_table)

    predictions_df = pd.DataFrame(predictions).transpose()
    st.write(predictions_df)
    all_predictions.append(predictions_df)  
        # ... (previous code for predictions_df)

        # Text Box 
    text_str = ''
    if benchmark_brands:
            for benchmark_brand in benchmark_brands:
                jklc_prices = [
                    district_df[f"JKLC ({week})"].iloc[0]
                    for week in week_names if f"JKLC ({week})" in district_df.columns
                ]
                benchmark_prices = [
                    district_df[f"{benchmark_brand} ({week})"].iloc[0]
                    for week in week_names
                    if f"{benchmark_brand} ({week})" in district_df.columns
                ]
                actual_diff = np.nan
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices) - 1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(
                                benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break

                brand_text = f"•Benchmark Brand: {benchmark_brand} → "
                brand_text += f"Actual Diff: {actual_diff:+.2f} Rs.|"
                if benchmark_brand in desired_diff and desired_diff[
                        benchmark_brand] is not None:
                    brand_desired_diff = desired_diff[benchmark_brand]
                    brand_text += f"Desired Diff: {brand_desired_diff:+.2f} Rs.| "
                    required_increase_decrease = brand_desired_diff - actual_diff
                    brand_text += f"Required Increase/Decrease in Price: {required_increase_decrease:+.2f} Rs."

                text_str += brand_text + "\n"

    ax.text(0.5,
                 -0.3,
                 text_str,
                 weight='bold',
                 ha='center',
                 va='center',
                 transform=ax.transAxes,
                 bbox=dict(facecolor='white',edgecolor='black',
                           boxstyle='round,pad=0.5'))
    plt.tight_layout()

        # Display plot in Streamlit
    st.pyplot(fig)

        # Download plot option
    buf = BytesIO()
    fig.savefig(buf, format='png', bbox_inches='tight')
    buf.seek(0)
    b64_data = base64.b64encode(buf.getvalue()).decode()
    href = f'<a href="data:image/png;base64,{b64_data}" download="district_plot_{district_name}.png">Download Plot as PNG</a>'
    st.markdown(href, unsafe_allow_html=True)

    # Download options for stats and predictions
    if st.button("Download Stats"):
        all_stats_df = pd.concat(all_stats_tables, keys=district_names)
        st.download_button("Download Stats (CSV)", all_stats_df.to_csv(), "stats.csv", "text/csv")
        st.download_button("Download Stats (Excel)", all_stats_df.to_excel(), "stats.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    if st.button("Download Predictions"):
        all_predictions_df = pd.concat(all_predictions, keys=district_names)
        st.download_button("Download Predictions (CSV)", all_predictions_df.to_csv(), "predictions.csv", "text/csv")
        st.download_button("Download Predictions (Excel)", all_predictions_df.to_excel(), "predictions.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# --- Streamlit App ---

st.title("Brand Price Trend Analysis")

uploaded_file = st.file_uploader("Upload Excel File", type="xlsx")

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file, skiprows=2)
        df = transform_data(df)

        zone_names = df["Zone"].unique().tolist()
        selected_zone = st.selectbox("Select Zone", zone_names)

        filtered_df = df[df["Zone"] == selected_zone]
        region_names = filtered_df["REGION"].unique().tolist()
        selected_region = st.selectbox("Select Region", region_names)

        filtered_df = df[df["REGION"] == selected_region]
        district_names = filtered_df["Dist Name"].unique().tolist()
        selected_districts = st.multiselect("Select District", district_names)

        if selected_districts:
            all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
            selected_benchmark_brands = st.multiselect("Select Benchmark Brands", benchmark_brands)

            desired_diff = {}
            if selected_benchmark_brands:
                for benchmark_brand in selected_benchmark_brands:
                    desired_diff[benchmark_brand] = st.number_input(f"Desired Diff for {benchmark_brand}", value=0.0, step=0.1)

            if st.button("Generate Plot"):
                plot_district_graph(df, selected_districts, selected_benchmark_brands, desired_diff)

    except Exception as e:
        st.error(f"Error processing file: {e}")
 
