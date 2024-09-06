import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from tqdm import tqdm
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
from scipy import stats
import matplotlib.backends.backend_pdf

# Function to transform data
def transform_data(df, week_names_input):
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
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Delhi', 'Haryana', 'Punjab'], 'North-I')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Uttar Pradesh(West)','Uttarakhand'], 'North-II')
    
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('EZ_East Zone', 'East Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('CZ_Central Zone', 'Central Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('NZ_North Zone', 'North Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('UPEZ_UP East Zone', 'UP East Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('upWZ_up West Zone', 'UP West Zone')
    transformed_df['Zone'] = transformed_df['Zone'].str.replace('WZ_West Zone', 'West Zone')
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    for i in tqdm(range(num_weeks), desc='Transforming data'):
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

# Function to plot district graph
def plot_district_graph(df, district_names, benchmark_brands, desired_diff, week_names, download_stats=False, download_predictions=False, download_pdf=False, diff_week=1):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    
    all_stats_table = []
    all_predictions = []
    if download_pdf:
        pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf")
    for i, district_name in enumerate(district_names):
        plt.figure(figsize=(10, 8))
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        stats_table_data = {}
        predictions = {}

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
                train_labels= np.array(valid_prices)
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
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        plt.ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        # Add region name above the title only for the first district
        if i == 0:
            plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)  # Added region name using plt.text
            plt.title(f"{district_name} - Brands Price Trend", weight='bold') # Keep the original title without region name
        else:
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()
        if stats_table_data:
            stats_table = pd.DataFrame(stats_table_data).transpose().round(2)
        else:
            stats_table = pd.DataFrame()  # Create an empty DataFrame if stats_table_data is empty
        st.write(stats_table)
        all_stats_table.append(stats_table)
        if predictions:
            predictions_df = pd.DataFrame(predictions).transpose()
        else:
            predictions_df = pd.DataFrame()  # Create an empty DataFrame if predictions is empty
        st.write(predictions_df)
        all_predictions.append(predictions_df)
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
        plt.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=plt.gca().transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
        plt.subplots_adjust(bottom=0.25)
        if download_pdf:
            pdf.savefig()
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        b64_data = base64.b64encode(buf.getvalue()).decode()
        st.markdown(f'<a download="district_plot_{district_name}.png" href="data:image/png;base64,{b64_data}">Download Plot as PNG</a>', unsafe_allow_html=True)
        st.pyplot()
    if download_pdf:
       pdf.close()
       with open("district_plots.pdf", "rb") as f:
           pdf_data = f.read()
       b64_pdf = base64.b64encode(pdf_data).decode()
       st.markdown(f'<a download="{region_name}.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>', unsafe_allow_html=True)   
    if download_stats:
        all_stats_df = pd.concat(all_stats_table, keys=district_names)
        for district_name in district_names:
            district_stats_df = all_stats_df.loc[district_name]
            stats_excel_path = f'stats_{district_name}.xlsx'
            district_stats_df.to_excel(stats_excel_path)
            excel_data = open(stats_excel_path, "rb").read()
            b64 = base64.b64encode(excel_data)
            payload = b64.decode        
        st.markdown(f'<a download="{district_name}_stats.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">Download {district_name} Stats as Excel File</a>', unsafe_allow_html=True)
    if download_predictions:
        all_predictions_df = pd.concat(all_predictions, keys=district_names)
        for district_name in district_names:
            district_predictions_df = all_predictions_df.loc[district_name]
            predictions_excel_path = f'predictions_{district_name}.xlsx'
            district_predictions_df.to_excel(predictions_excel_path)
            excel_data = open(predictions_excel_path, "rb").read()
            b64 = base64.b64encode(excel_data)
            payload = b64.decode()
        st.markdown(f'<a download="{district_name}_predictions.xlsx" href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}">Download {district_name} Predictions as Excel File</a>', unsafe_allow_html=True)

# Load the data
df = pd.read_excel("weekwise_data.xlsx")

# Get the week names from user input
week_names_input = st.text_input("Enter week names separated by comma", "Week 1, Week 2, Week 3").split(", ")

# Transform data based on week names input
transformed_df = transform_data(df, week_names_input)

# Get unique district names
district_names = transformed_df['Dist Name'].unique()

# Select districts and benchmark brands for plotting
selected_districts = st.multiselect("Select districts", district_names)
benchmark_brands = st.multiselect("Select benchmark brands", brands)

# Get the desired price difference from user input
desired_diff = {brand: st.number_input(f"Enter desired price difference for {brand} (in Rs.)", value=1000) for brand in benchmark_brands}

# Plot district graphs based on user input
plot_district_graph(transformed_df, selected_districts, benchmark_brands, desired_diff, week_names_input, download_stats=True, download_predictions=True, download_pdf=True)

