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

def inspect_excel(file):
    # Read the first few rows of the Excel file
    df = pd.read_excel(file, header=None, nrows=5)
    return df

def transform_data(df):
    # Identify the row with brand names (usually the second row)
    brand_row = df.iloc[1]
    brands = [brand for brand in brand_row.dropna().unique() if isinstance(brand, str)]

    # Extract week/month names from the first row
    week_month_row = df.iloc[0]
    week_month_names = [name for name in week_month_row.dropna().unique() if isinstance(name, str) and name not in brands]

    # Identify the metadata columns
    metadata_columns = df.columns[:4].tolist()

    # Create the transformed dataframe
    transformed_df = df.iloc[2:].copy()  # Start from the third row
    transformed_df.columns = metadata_columns + [f"{brand} ({week_month})" for week_month in week_month_names for brand in brands]
    
    # Convert price columns to numeric, replacing any non-numeric values with NaN
    price_columns = transformed_df.columns[4:]  # Assuming first 4 columns are metadata
    transformed_df[price_columns] = transformed_df[price_columns].apply(pd.to_numeric, errors='coerce')

    return transformed_df, week_month_names, brands, metadata_columns

def plot_district_graph(df, district_name, benchmark_brands, desired_diff, selected_week_month, brands, metadata_columns):
    fig, (ax, ax2) = plt.subplots(1, 2, figsize=(20, 10), gridspec_kw={'width_ratios': [3, 1]})
    district_df = df[df[metadata_columns[3]] == district_name]  # Assuming Dist Name is the 4th column
    
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

def calculate_stats_and_predictions(df, district_name, brands, metadata_columns):
    district_df = df[df[metadata_columns[3]] == district_name]  # Assuming Dist Name is the 4th column
    
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
            # Inspect the Excel file
            inspect_df = inspect_excel(uploaded_file)
            st.write("Preview of the first few rows:")
            st.dataframe(inspect_df)

            # Ask user to confirm the structure
            if st.button("Confirm and Process Data"):
                st.session_state.df, week_month_names, brands, metadata_columns = transform_data(inspect_df)
                st.success("File processed successfully!")
                
                # Display the processed dataframe
                st.write("Processed Data Preview:")
                st.dataframe(st.session_state.df.head())

        except Exception as e:
            st.error(f"Error reading file: {str(e)}. Please ensure it is a valid Excel file.")
            return

    if st.session_state.df is not None:
        df = st.session_state.df
        
        col1, col2, col3 = st.columns(3)
        with col1:
            zone = st.selectbox("Select Zone", options=df[metadata_columns[0]].unique())
        with col2:
            region = st.selectbox("Select Region", options=df[df[metadata_columns[0]] == zone][metadata_columns[1]].unique())
        with col3:
            selected_week_month = st.selectbox("Select Week/Month for Plot", options=week_month_names)

        districts = df[(df[metadata_columns[0]] == zone) & (df[metadata_columns[1]] == region)][metadata_columns[3]].unique()
        selected_districts = st.multiselect("Select Districts", options=districts)

        benchmark_brands = brands[1:]  # Exclude JKLC as it's the reference brand

        # Create a form for each selected district
        for district in selected_districts:
            with st.expander(f"Set benchmark brands for {district}"):
                st.session_state.district_benchmarks[district] = {}
                selected_benchmarks = st.multiselect(f"Select Benchmark Brands for {district}", options=benchmark_brands)
                for brand in selected_benchmarks:
                    st.session_state.district_benchmarks[district][brand] = st.number_input(f"Desired Diff for {brand} in {district}", value=0.0)

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
                        selected_week_month,
                        brands,
                        metadata_columns
                    )
                    st.pyplot(fig)
                    
                    # Calculate stats and predictions based on all dates
                    stats, predictions = calculate_stats_and_predictions(df, district, brands, metadata_columns)
                    
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
