import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import xgboost as xgb
from scipy import stats
from io import BytesIO
import base64

# Set page config
st.set_page_config(page_title="Brand Price Analysis", layout="wide")

@st.cache_data
def load_and_transform_data(uploaded_file):
    df = pd.read_excel(uploaded_file, skiprows=2)
    return transform_data(df)

def transform_data(df):
    # [The transform_data function remains the same as before]
    # ...

@st.cache_data
def precompute_data(df, brands, week_names):
    all_data = {}
    for district_name in df["Dist Name"].unique():
        district_df = df[df["Dist Name"] == district_name]
        district_data = {}
        for brand in brands:
            brand_prices = [district_df[f"{brand} ({week})"].iloc[0] if f"{brand} ({week})" in district_df.columns else np.nan for week in week_names]
            valid_prices = [p for p in brand_prices if not np.isnan(p)]
            
            if valid_prices:
                price_diff = valid_prices[-1] - valid_prices[0]
                stats = calculate_stats(valid_prices)
                prediction = make_prediction(valid_prices)
            else:
                price_diff = np.nan
                stats = {stat: np.nan for stat in ['Min', 'Max', 'Average', 'Median', 'First Quartile', 'Third Quartile', 'Variance', 'Skewness', 'Kurtosis']}
                prediction = {'Prediction': np.nan, 'Confidence Interval': (np.nan, np.nan)}
            
            district_data[brand] = {
                'prices': brand_prices,
                'price_diff': price_diff,
                'stats': stats,
                'prediction': prediction
            }
        all_data[district_name] = district_data
    return all_data

def calculate_stats(valid_prices):
    return {
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

def make_prediction(valid_prices):
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

        return {
            'Prediction': prediction[0],
            'Confidence Interval': confidence_interval[0]
        }
    else:
        return {
            'Prediction': np.nan,
            'Confidence Interval': (np.nan, np.nan)
        }

def plot_district_graph(district_name, district_data, week_names, benchmark_brands, desired_diff):
    fig, ax = plt.subplots(figsize=(12, 10))
    brands = list(district_data.keys())

    for brand in brands:
        brand_data = district_data[brand]
        ax.plot(week_names, brand_data['prices'], marker='o', linestyle='-', label=f"{brand} ({brand_data['price_diff']:.0f})")

        for week, price in zip(week_names, brand_data['prices']):
            if not np.isnan(price):
                ax.text(week, price, str(round(price)), fontsize=8)

    ax.set_xlabel('Month/Week', weight='bold')
    ax.set_ylabel('Whole Sale Price (in Rs.)', weight='bold')
    ax.set_title(f"{district_name} - Brands Price Trend", weight='bold')
    plt.xticks(rotation=45)
    plt.tight_layout()

    ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, prop={'weight': 'bold'})

    text_str = ''
    if benchmark_brands:
        for benchmark_brand in benchmark_brands:
            jklc_prices = district_data['JKLC']['prices']
            benchmark_prices = district_data[benchmark_brand]['prices']
            actual_diff = np.nan
            if jklc_prices and benchmark_prices:
                for i in range(len(jklc_prices) - 1, -1, -1):
                    if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                        actual_diff = jklc_prices[i] - benchmark_prices[i]
                        break

            brand_text = f"•Benchmark Brand: {benchmark_brand} → "
            brand_text += f"Actual Diff: {actual_diff:+.2f} Rs.|"
            if benchmark_brand in desired_diff and desired_diff[benchmark_brand] is not None:
                brand_desired_diff = desired_diff[benchmark_brand]
                brand_text += f"Desired Diff: {brand_desired_diff:+.2f} Rs.| "
                required_increase_decrease = brand_desired_diff - actual_diff
                brand_text += f"Required Increase/Decrease in Price: {required_increase_decrease:+.2f} Rs."

            text_str += brand_text + "\n"

    plt.figtext(0.5, 0.01, text_str, ha='center', va='center', bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
    plt.subplots_adjust(bottom=0.3)

    return fig

def main():
    st.title("Brand Price Analysis")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            df = load_and_transform_data(uploaded_file)
            st.success("File uploaded successfully!")
        except Exception as e:
            st.error(f"Error reading file: {e}. Please ensure it is a valid Excel file.")
            return

        brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
        week_names = sorted(list(set([col.split(' (')[1].split(')')[0] for col in df.columns if '(' in col])),
                            key=lambda x: ('June July August September October November December January February March April May'.split().index(x.split()[-1]), int(x.split('-')[1]) if '-' in x else 0))

        all_data = precompute_data(df, brands, week_names)

        col1, col2 = st.columns(2)
        with col1:
            zone = st.selectbox("Select Zone", options=df["Zone"].unique())
        with col2:
            region = st.selectbox("Select Region", options=df[df["Zone"] == zone]["REGION"].unique())

        districts = df[(df["Zone"] == zone) & (df["REGION"] == region)]["Dist Name"].unique()
        selected_districts = st.multiselect("Select Districts", options=districts)

        benchmark_brands = ['UTCL', 'JKS', 'Ambuja', 'Wonder', 'Shree']
        selected_benchmark_brands = st.multiselect("Select Benchmark Brands", options=benchmark_brands)

        desired_diff = {}
        for brand in selected_benchmark_brands:
            desired_diff[brand] = st.number_input(f"Desired Diff for {brand}", value=0)

        if st.button("Generate Analysis"):
            if selected_districts and selected_benchmark_brands:
                for district_name in selected_districts:
                    district_data = all_data[district_name]
                    fig = plot_district_graph(district_name, district_data, week_names, selected_benchmark_brands, desired_diff)
                    st.pyplot(fig)

                    st.write(f"### Statistics for {district_name}")
                    stats_df = pd.DataFrame({brand: data['stats'] for brand, data in district_data.items()}).transpose()
                    st.dataframe(stats_df.round(2))

                    st.write(f"### Predictions for {district_name}")
                    predictions_df = pd.DataFrame({brand: data['prediction'] for brand, data in district_data.items()}).transpose()
                    st.dataframe(predictions_df)
            else:
                st.warning("Please select at least one district and one benchmark brand.")

if __name__ == "__main__":
    main()
