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

# Global variables
df = None
desired_diff_input = {}
week_names = []

def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    # Region name replacements
    region_replacements = {
        '12_Madhya Pradesh(west)': 'Madhya Pradesh(West)',
        '20_Rajasthan': 'Rajasthan', '50_Rajasthan III': 'Rajasthan', '80_Rajasthan II': 'Rajasthan',
        '33_Chhattisgarh(2)': 'Chhattisgarh', '38_Chhattisgarh(3)': 'Chhattisgarh', '39_Chhattisgarh(1)': 'Chhattisgarh',
        '07_Haryana 1': 'Haryana', '07_Haryana 2': 'Haryana',
        '06_Gujarat 1': 'Gujarat', '66_Gujarat 2': 'Gujarat', '67_Gujarat 3': 'Gujarat', '68_Gujarat 4': 'Gujarat', '69_Gujarat 5': 'Gujarat',
        '13_Maharashtra': 'Maharashtra(West)',
        '24_Uttar Pradesh': 'Uttar Pradesh(West)',
        '35_Uttarakhand': 'Uttarakhand',
        '83_UP East Varanasi Region': 'Varanasi',
        '83_UP East Lucknow Region': 'Lucknow',
        '30_Delhi': 'Delhi',
        '19_Punjab': 'Punjab',
        '09_Jammu&Kashmir': 'Jammu&Kashmir',
        '08_Himachal Pradesh': 'Himachal Pradesh',
        '82_Maharashtra(East)': 'Maharashtra(East)',
        '81_Madhya Pradesh': 'Madhya Pradesh(East)',
        '34_Jharkhand': 'Jharkhand',
        '18_ODISHA': 'Odisha',
        '04_Bihar': 'Bihar',
        '27_Chandigarh': 'Chandigarh',
        '82_Maharashtra (East)': 'Maharashtra(East)',
        '25_West Bengal': 'West Bengal'
    }
    
    transformed_df['REGION'] = transformed_df['REGION'].replace(region_replacements)
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Delhi', 'Haryana', 'Punjab'], 'North-I')
    
    # Zone name replacements
    zone_replacements = {
        'EZ_East Zone': 'East Zone',
        'CZ_Central Zone': 'Central Zone',
        'NZ_North Zone': 'North Zone',
        'UPEZ_UP East Zone': 'UP East Zone',
        'upWZ_up West Zone': 'UP West Zone',
        'WZ_West Zone': 'West Zone'
    }
    
    transformed_df['Zone'] = transformed_df['Zone'].replace(zone_replacements)
    
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True)
    
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands, desired_diff, week_names, diff_week=1):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    
    all_stats_table = []
    all_predictions = []
    
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
            if len(valid_prices) > diff_week:
                price_diff = valid_prices[-1] - valid_prices[diff_week]
            else:
                price_diff = np.nan
            price_diffs.append(price_diff)
            line, = ax.plot(week_names, brand_prices, marker='o', linestyle='-', label=f"{brand} ({price_diff:.0f})")
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
                stats_table_data[brand] = {key: np.nan for key in ['Min', 'Max', 'Average', 'Median', 'First Quartile', 'Third Quartile', 'Variance', 'Skewness', 'Kurtosis']}
            
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
        plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()
        
        if stats_table_data:
            stats_table = pd.DataFrame(stats_table_data).transpose().round(2)
        else:
            stats_table = pd.DataFrame()
        all_stats_table.append(stats_table)
        
        if predictions:
            predictions_df = pd.DataFrame(predictions).transpose()
        else:
            predictions_df = pd.DataFrame()
        all_predictions.append(predictions_df)
        
        text_str = ''
        if benchmark_brands:
            brand_texts = []
            max_left_length = 0
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
        
        plt.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=ax.transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
        plt.subplots_adjust(bottom=0.25)
        
        st.pyplot(fig)
        st.write(f"Stats for {district_name}:")
        st.dataframe(stats_table)
        st.write(f"Predictions for {district_name}:")
        st.dataframe(predictions_df)
    
    return all_stats_table, all_predictions

def main():
    st.title("Data Analysis and Visualization App")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file, skiprows=2)
        
        brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
        brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
        num_weeks = len(brand_columns) // 6
        
        global week_names
        week_names = []
        for i in range(num_weeks):
            week_name = st.text_input(f"Week {i+1} name:", f"Week {i+1}")
            week_names.append(week_name)
        
        if st.button("Process Data"):
            df_transformed = transform_data(df, week_names)
            
            zone_names = df_transformed["Zone"].unique().tolist()
            selected_zone = st.selectbox("Select Zone", zone_names)
            
            filtered_df = df_transformed[df_transformed["Zone"] == selected_zone]
            region_names = filtered_df["REGION"].unique().tolist()
            selected_region = st.selectbox("Select Region", region_names)
            
            filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
            district_names = filtered_df["Dist Name"].unique().tolist()
            selected_districts = st.multiselect("Select Districts", district_names)
            
            all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
            selected_benchmark_brands = st.multiselect("Select Benchmark Brands", benchmark_brands)
            
            desired_diff = {}
            for brand in selected_benchmark_brands:
                desired_diff[brand] = st.number_input(f"Desired Diff for {brand}:", value=0)
            
            diff_week = st.slider("Diff Week", min_value=0, max_value=num_weeks-1, value=1)
            
            if st.button("Generate Plots"):
                all_stats, all_predictions = plot_district_graph(filtered_df, selected_districts, selected_benchmark_brands, desired_diff, week_names, diff_week)
                
                # Offer downloads
                if st.button("Download Stats"):
                    for district, stats in zip(selected_districts, all_stats):
                        csv = stats.to_csv(index=True)
                        b64 = base64.b64encode(csv.encode()).decode()
                        href = f'<a href="data:file/csv;base64,{b64}" download="stats_{district}.csv">Download {district} Stats CSV</a>'
                        st.markdown(href, unsafe_allow_html=True)

                if st.button("Download Predictions"):
                    for district, predictions in zip(selected_districts, all_predictions):
                        csv = predictions.to_csv(index=True)
                        b64 = base64.b64encode(csv.encode()).decode()
                        href = f'<a href="data:file/csv;base64,{b64}" download="predictions_{district}.csv">Download {district} Predictions CSV</a>'
                        st.markdown(href, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
