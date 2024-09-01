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
if 'district_benchmarks' not in st.session_state:
    st.session_state.district_benchmarks = {}

def transform_data(df):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    # Extract week/month names from the first row
    week_month_names = df.iloc[0].dropna().tolist()[4:]  # Skip the first 4 columns
    
    # Process brand columns
    for i, week_month in enumerate(week_month_names):
        start_idx = 4 + i * len(brands)
        end_idx = 4 + (i + 1) * len(brands)
        week_data = df.iloc[2:, start_idx:end_idx].copy()  # Start from the third row
        week_data.columns = [f"{brand} ({week_month})" for brand in brands]
        week_data = week_data.apply(pd.to_numeric, errors='coerce')
        transformed_df = pd.concat([transformed_df, week_data], axis=1)
    
    return transformed_df, week_month_names

def plot_district_graph(df, district_name, benchmark_brands, desired_diff, selected_week_month):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(20, 10), gridspec_kw={'width_ratios': [3, 1]})
    district_df = df[df["Dist Name"] == district_name]
    
    for brand in brands:
        brand_price = district_df[f"{brand} ({selected_week_month})"].iloc[0]
        ax.bar(brand, brand_price)
        ax.text(brand, brand_price, f'{brand_price:.0f}', ha='center', va='bottom')

    ax.set_ylabel('Whole Sale Price (in Rs.)', weight='bold')
    ax.set_title(f"{district_name} - Brands Price for {selected_week_month}", weight='bold')
    plt.setp(ax.get_xticklabels(), rotation=45)

    # Benchmark brand information on the right side
    ax2.axis('off')
    text_str = 'Benchmark Brands:\n\n'
    if benchmark_brands:
        for benchmark_brand in benchmark_brands:
            jklc_price = district_df[f"JKLC ({selected_week_month})"].iloc[0]
            benchmark_price = district_df[f"{benchmark_brand} ({selected_week_month})"].iloc[0]
            actual_diff = jklc_price - benchmark_price

            brand_text = f"╔══ {benchmark_brand} ══╗\n"
            brand_text += f"║ Actual Diff: {actual_diff:+.2f} Rs. ║\n"
            if benchmark_brand in desired_diff and desired_diff[benchmark_brand] is not None:
                brand_desired_diff = desired_diff[benchmark_brand]
                brand_text += f"║ Desired Diff: {brand_desired_diff:+.2f} Rs. ║\n"
                required_increase_decrease = brand_desired_diff - actual_diff
                brand_text += f"║ Required Change: {required_increase_decrease:+.2f} Rs. ║\n"
            brand_text += "╚" + "═" * (len(benchmark_brand) + 6) + "╝\n\n"

            text_str += brand_text

    ax2.text(0, 0.95, text_str, va='top', ha='left', transform=ax2.transAxes, family='monospace', fontsize=10)

    plt.tight_layout()
    return fig

def calculate_stats_and_predictions(df, district_name):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    district_df = df[df["Dist Name"] == district_name]
    
    stats_table_data = {}
    predictions = {}
    
    for brand in brands:
        brand_prices = district_df[[col for col in district_df.columns if brand in col]].values.flatten()
        valid_prices = brand_prices[~np.isnan(brand_prices)]
        
        if len(valid_prices) > 0:
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
            
            if len(valid_prices) > 2:
                train_data = np.array(range(len(valid_prices))).reshape(-1, 1)
                model = xgb.XGBRegressor(objective='reg:squarederror')
                model.fit(train_data, valid_prices)
                next_week = len(valid_prices)
                prediction = model.predict(np.array([[next_week]]))
                
                errors = abs(model.predict(train_data) - valid_prices)
                confidence = 0.95
                n = len(valid_prices)
                t_crit = stats.t.ppf((1 + confidence) / 2, n - 1)
                margin_of_error = t_crit * errors.std() / np.sqrt(n)
                confidence_interval = (prediction - margin_of_error, prediction + margin_of_error)
                
                predictions[brand] = {
                    'Prediction': prediction[0],
                    'Confidence Interval': confidence_interval
                }
            else:
                predictions[brand] = {
                    'Prediction': np.nan,
                    'Confidence Interval': (np.nan, np.nan)
                }
        else:
            stats_table_data[brand] = {stat: np.nan for stat in ['Min', 'Max', 'Average', 'Median', 'First Quartile', 'Third Quartile', 'Variance', 'Skewness', 'Kurtosis']}
            predictions[brand] = {
                'Prediction': np.nan,
                'Confidence Interval': (np.nan, np.nan)
            }
    
    return stats_table_data, predictions

def main():
    st.title("Brand Price Analysis")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            st.session_state.df, week_month_names = transform_data(df)
            st.success("File uploaded successfully!")
        except Exception as e:
            st.error(f"Error reading file: {e}. Please ensure it is a valid Excel file.")
            return

    if st.session_state.df is not None:
        df = st.session_state.df
        
        col1, col2, col3 = st.columns(3)
        with col1:
            zone = st.selectbox("Select Zone", options=df["Zone"].unique())
        with col2:
            region = st.selectbox("Select Region", options=df[df["Zone"] == zone]["REGION"].unique())
        with col3:
            selected_week_month = st.selectbox("Select Week/Month for Plot", options=week_month_names)

        districts = df[(df["Zone"] == zone) & (df["REGION"] == region)]["Dist Name"].unique()
        selected_districts = st.multiselect("Select Districts", options=districts)

        benchmark_brands = ['UTCL', 'JKS', 'Ambuja', 'Wonder', 'Shree']

        # Create a form for each selected district
        for district in selected_districts:
            with st.expander(f"Set benchmark brands for {district}"):
                st.session_state.district_benchmarks[district] = {}
                selected_benchmarks = st.multiselect(f"Select Benchmark Brands for {district}", options=benchmark_brands)
                for brand in selected_benchmarks:
                    st.session_state.district_benchmarks[district][brand] = st.number_input(f"Desired Diff for {brand} in {district}", value=0)

        if st.button("Generate Analysis"):
            if selected_districts:
                for district in selected_districts:
                    st.write(f"### Analysis for {district}")
                    
                    # Generate plot for selected week/month
                    fig = plot_district_graph(
                        df, 
                        district, 
                        st.session_state.district_benchmarks[district].keys(),
                        st.session_state.district_benchmarks[district],
                        selected_week_month
                    )
                    st.pyplot(fig)
                    
                    # Calculate stats and predictions based on all dates
                    stats, predictions = calculate_stats_and_predictions(df, district)
                    
                    st.write("#### Descriptive Statistics (based on all dates)")
                    st.dataframe(pd.DataFrame(stats).transpose().round(2))
                    
                    st.write("#### Predictions (based on all dates)")
                    st.dataframe(pd.DataFrame(predictions).transpose())
                    
                    # Add download buttons for plot
                    buf = BytesIO()
                    fig.savefig(buf, format="png")
                    buf.seek(0)
                    st.download_button(
                        label=f"Download {district} Plot",
                        data=buf,
                        file_name=f"{district}_{selected_week_month}_plot.png",
                        mime="image/png"
                    )
            else:
                st.warning("Please select at least one district.")

if __name__ == "__main__":
    main()
