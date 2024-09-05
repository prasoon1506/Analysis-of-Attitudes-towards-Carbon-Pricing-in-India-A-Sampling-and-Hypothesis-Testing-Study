import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
from tqdm import tqdm
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
from scipy import stats
import matplotlib.backends.backend_pdf

def transform_data(df):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('12_Madhya Pradesh(west)', 'Madhya Pradesh(West)')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['20_Rajasthan', '50_Rajasthan III', '80_Rajasthan II'], 'Rajasthan')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['33_Chhattisgarh(2)', '38_Chhattisgarh(3)', '39_Chhattisgarh(1)'], 'Chhattisgarh')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['07_Haryana 1', '07_Haryana 2'], 'Haryana')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['06_Gujarat 1', '66_Gujarat 2', '67_Gujarat 3','68_Gujarat 4','69_Gujarat 5'], 'Gujarat')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('13_Maharashtra', 'Maharashtra(West)')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('24_Uttar Pradesh', 'Uttar Pradesh(West)')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('35_Uttarakhand', 'Uttarakhand')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('83_UP East Varanasi Region', 'Varanasi')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('83_UP East Lucknow Region', 'Lucknow')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('30_Delhi', 'Delhi')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('19_Punjab', 'Punjab')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('09_Jammu&Kashmir', 'Jammu&Kashmir')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('08_Himachal Pradesh', 'Himachal Pradesh')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('82_Maharashtra(East)', 'Maharashtra(East)')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('81_Madhya Pradesh', 'Madhya Pradesh(East)')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('34_Jharkhand', 'Jharkhand')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('18_ODISHA', 'Odisha')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('04_Bihar', 'Bihar')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('27_Chandigarh', 'Chandigarh')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('82_Maharashtra (East)', 'Maharashtra(East)')
    transformed_df['REGION'] = transformed_df['REGION'].str.replace('25_West Bengal', 'West Bengal')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('EZ_East Zone', 'East Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('CZ_Central Zone', 'Central Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('NZ_North Zone', 'North Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('UPEZ_UP East Zone', 'UP East Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('upWZ_up West Zone', 'UP West Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('WZ_West Zone', 'West Zone')
    
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
        if i == 0:
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
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df,
                                  week_data,
                                  left_index=True,
                                  right_index=True)
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands, desired_diff):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    week_names = list(
        set([col.split(' (')[1].split(')')[0] for col in df.columns
             if '(' in col]))

    def sort_week_names(week_name):
        if ' ' in week_name:
            week, month = week_name.split()
            week_num = int(week.split('-')[1])
        else:
            week_num = 0
            month = week_name
        month_order = [
            'June', 'July', 'August', 'September', 'October', 'November',
            'December', 'January', 'February', 'March', 'April', 'May'
        ]
        month_num = month_order.index(month)
        return month_num * 10 + week_num

    week_names.sort(key=sort_week_names)

    for district_name in district_names:
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
            if len(valid_prices) > 2:
                train_data = np.array(range(len(valid_prices))).reshape(-1, 1)
                train_labels = np.array(valid_prices)
                model = xgb.XGBRegressor(objective='reg:squarederror')
                model.fit(train_data, train_labels)
                next_week = len(valid_prices)
                prediction = model.predict(np.array([[next_week]]))
                errors = abs(model.predict(train_data) - train_labels)
                confidence = 0.95
                n = len(valid_prices)
                t_crit = stats.t.ppf((1 + confidence) / 2, n - 1)
                margin_of_error = t_crit * errors.std() / np.sqrt(n)
                confidence_interval = (prediction - margin_of_error, prediction + margin_of_error)
                predictions[brand] = {'Prediction': prediction[0], 'Confidence Interval': confidence_interval}
            else:
                predictions[brand] = {'Prediction': np.nan, 'Confidence Interval': (np.nan, np.nan)}

        ax.grid(False)
        ax.set_xlabel('Month/Week', weight='bold')
        ax.set_ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=ax.transAxes, weight='bold', fontsize=16)
        ax.set_title(f"{district_name} - Brands Price Trend", weight='bold')
        
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()

        st.pyplot(fig)

        stats_table = pd.DataFrame(stats_table_data).transpose().round(2)
        st.write("Statistics:")
        st.dataframe(stats_table)

        predictions_df = pd.DataFrame(predictions).transpose()
        st.write("Predictions:")
        st.dataframe(predictions_df)

        if benchmark_brands:
            for benchmark_brand in benchmark_brands:
                jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
                benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
                actual_diff = np.nan
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices) - 1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break
                desired_diff_str = f" ({desired_diff[benchmark_brand]:.0f} Rs.)" if benchmark_brand in desired_diff else ""
                st.write(f"Benchmark Brand: {benchmark_brand}{desired_diff_str}")
                st.write(f"Actual Diff: {actual_diff:+.2f} Rs.")

def main():
    st.title("Brand Price Analysis")

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
            selected_districts = st.multiselect("Select District(s)", district_names)

            all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
            selected_benchmark_brands = st.multiselect("Select Benchmark Brands", benchmark_brands)

            desired_diff = {}
            for brand in selected_benchmark_brands:
                desired_diff[brand] = st.number_input(f"Desired Diff for {brand}", value=0)

            if st.button("Generate Plots"):
                plot_district_graph(df, selected_districts, selected_benchmark_brands, desired_diff)

            # Add download buttons
            if st.button("Download Statistics"):
                stats_csv = BytesIO()
                stats_table.to_csv(stats_csv, index=True)
                stats_csv.seek(0)
                st.download_button(
                    label="Download Statistics CSV",
                    data=stats_csv,
                    file_name="statistics.csv",
                    mime="text/csv"
                )

            if st.button("Download Predictions"):
                predictions_csv = BytesIO()
                predictions_df.to_csv(predictions_csv, index=True)
                predictions_csv.seek(0)
                st.download_button(
                    label="Download Predictions CSV",
                    data=stats_csv,
                    file_name="Predictions.csv",
                    mime="text/csv"
                )
        except Exception as e:
            st.error(f"Error reading file: {e}. Please ensure it is a valid Excel file.")

if __name__ == "__main__":
    main()
