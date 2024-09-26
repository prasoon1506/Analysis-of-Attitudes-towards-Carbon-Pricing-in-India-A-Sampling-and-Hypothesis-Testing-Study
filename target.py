import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy import stats

# Load data
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

# Preprocess data
def preprocess_data(df_row):
    processed_data = {}
    
    # Year-over-Year Growth
    processed_data['YoY_Growth'] = (df_row['Monthly Achievement(Aug)'] - df_row['Total Aug 2023']) / df_row['Total Aug 2023']
    
    # Monthly Growth Rates
    months = ['Apr', 'May', 'June', 'July', 'Aug']
    for i in range(1, len(months)):
        processed_data[f'Growth_{months[i-1]}_{months[i]}'] = (df_row[f'Monthly Achievement({months[i]})'] - df_row[f'Monthly Achievement({months[i-1]})']) / df_row[f'Monthly Achievement({months[i-1]})']
    
    # Average Monthly Growth
    processed_data['Avg_Monthly_Growth'] = np.mean([processed_data[f'Growth_{months[i-1]}_{months[i]}'] for i in range(1, len(months))])
    
    # Seasonal Index
    processed_data['Seasonal_Index'] = df_row['Monthly Achievement(Aug)'] / df_row['Total Sep 2023']
    
    return processed_data

# Prediction function
def predict_sales(df, region, brand):
    df_filtered = df[(df['Zone'] == region) & (df['Brand'] == brand)]
    
    if len(df_filtered) == 0:
        return None, None, None
    
    row = df_filtered.iloc[0]
    processed_data = preprocess_data(row)
    
    # Simple time series forecast
    months = ['Apr', 'May', 'June', 'July', 'Aug']
    sales = [row[f'Monthly Achievement({month})'] for month in months]
    
    # Calculate trend
    x = np.arange(len(months))
    slope, intercept, r_value, p_value, std_err = stats.linregress(x, sales)
    trend = slope * (len(months)) + intercept
    
    # Apply seasonal adjustment
    seasonal_factor = processed_data['Seasonal_Index']
    
    # Apply growth rate
    growth_factor = 1 + processed_data['Avg_Monthly_Growth']
    
    # Final prediction
    sept_prediction = trend * seasonal_factor * growth_factor
    
    # Calculate prediction interval
    residuals = np.array(sales) - (slope * x + intercept)
    std_residuals = np.std(residuals)
    pi_range = 1.96 * std_residuals  # 95% prediction interval
    
    pi_lower = max(0, sept_prediction - pi_range)
    pi_upper = sept_prediction + pi_range
    
    # Use the prediction interval as a proxy for the confidence interval
    ci_lower, ci_upper = pi_lower, pi_upper
    
    return sept_prediction, (ci_lower, ci_upper), (pi_lower, pi_upper)

# Visualization function
def create_visualization(df_row, region, brand, prediction, ci, pi):
    fig, ax = plt.subplots(figsize=(12, 6))
    
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']
    achievements = [df_row[f'Monthly Achievement({m})'] for m in months[:-1]] + [prediction]
    targets = [df_row[f'Month Tgt ({m})'] for m in months]
    
    ax.bar(months, targets, alpha=0.5, label='Target')
    ax.bar(months, achievements, alpha=0.7, label='Achievement')
    
    ax.set_title(f'Monthly Targets and Achievements for {region} - {brand}')
    ax.set_xlabel('Month')
    ax.set_ylabel('Sales')
    ax.legend()
    
    ax.errorbar('Sep', prediction, yerr=[[prediction-ci[0]], [ci[1]-prediction]], 
                fmt='o', color='r', capsize=5, label='95% CI')
    ax.errorbar('Sep', prediction, yerr=[[prediction-pi[0]], [pi[1]-prediction]], 
                fmt='o', color='g', capsize=5, label='95% PI')
    
    ax.legend()
    
    return fig

# Streamlit app
def main():
    st.set_page_config(page_title="Sales Prediction App", page_icon="📊", layout="wide")
    
    st.title("📊 Sales Prediction App")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        
        st.sidebar.header("Filters")
        region = st.sidebar.selectbox("Select Region", df['Zone'].unique())
        brand = st.sidebar.selectbox("Select Brand", df['Brand'].unique())
        
        if st.sidebar.button("Generate Prediction"):
            with st.spinner("Generating prediction..."):
                prediction, ci, pi = predict_sales(df, region, brand)
            
            if prediction is not None:
                st.success("Prediction generated successfully!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Prediction Results")
                    st.write(f"Predicted September 2024 sales: {prediction:.2f}")
                    st.write(f"95% Confidence Interval: ({ci[0]:.2f}, {ci[1]:.2f})")
                    st.write(f"95% Prediction Interval: ({pi[0]:.2f}, {pi[1]:.2f})")
                
                with col2:
                    fig = create_visualization(df[(df['Zone'] == region) & (df['Brand'] == brand)].iloc[0], region, brand, prediction, ci, pi)
                    st.pyplot(fig)
                
                st.subheader("Interpretation")
                relative_ci_width = (ci[1] - ci[0]) / prediction * 100
                relative_pi_width = (pi[1] - pi[0]) / prediction * 100
                st.write(f"Relative Confidence Interval width: {relative_ci_width:.2f}%")
                st.write(f"Relative Prediction Interval width: {relative_pi_width:.2f}%")
                
                if relative_ci_width < 10:
                    st.write("The model's confidence interval is narrow, indicating high precision in the estimate.")
                elif relative_ci_width < 20:
                    st.write("The model's confidence interval is moderately narrow, indicating good precision in the estimate.")
                else:
                    st.write("The model's confidence interval is wide, indicating uncertainty in the estimate.")
            else:
                st.error("Unable to generate prediction. No data found for the selected region and brand.")
        
        st.sidebar.header("Prediction Techniques")
        st.sidebar.write("""
        1. Time Series Forecasting: Uses historical data to identify trends and seasonality.
        2. Trend Analysis: Calculates the overall direction of sales over time.
        3. Seasonal Adjustment: Accounts for recurring patterns in sales data.
        4. Growth Rate Application: Incorporates recent growth trends into the forecast.
        5. Confidence and Prediction Intervals: Provides a range of possible outcomes.
        """)

if __name__ == "__main__":
    main()
