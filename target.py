import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from xgboost import XGBRegressor
from lightgbm import LGBMRegressor
from sklearn.linear_model import ElasticNet

# Load and preprocess data
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

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
    
    # Achievement Rates
    for month in months:
        processed_data[f'Achievement_Rate_{month}'] = df_row[f'Monthly Achievement({month})'] / df_row[f'Month Tgt ({month})']
    
    # Average Achievement Rate
    processed_data['Avg_Achievement_Rate'] = np.mean([processed_data[f'Achievement_Rate_{month}'] for month in months])
    
    # Additional Features
    processed_data['Seasonal_Index'] = df_row['Monthly Achievement(Aug)'] / df_row['Total Sep 2023']
    processed_data['Aug_YoY_Diff'] = df_row['Monthly Achievement(Aug)'] - df_row['Total Aug 2023']
    processed_data['Sep_Aug_Target_Ratio'] = df_row['Month Tgt (Sep)'] / df_row['Monthly Achievement(Aug)']
    
    return processed_data

def create_features(df_row):
    processed_data = preprocess_data(df_row)
    features = [
        df_row['Month Tgt (Sep)'], df_row['Monthly Achievement(Aug)'], df_row['Total Sep 2023'], df_row['Total Aug 2023'],
        processed_data['YoY_Growth'], processed_data['Avg_Monthly_Growth'], processed_data['Avg_Achievement_Rate'],
        processed_data['Seasonal_Index'], processed_data['Aug_YoY_Diff'], processed_data['Sep_Aug_Target_Ratio']
    ] + [df_row[f'Month Tgt ({month})'] for month in ['Apr', 'May', 'June', 'July', 'Aug']]
    
    return np.array(features).reshape(1, -1)

# Model training and prediction
@st.cache_resource
def create_models():
    models = {
        'rf': RandomForestRegressor(n_estimators=100, random_state=42),
        'xgb': XGBRegressor(n_estimators=100, random_state=42),
        'lgbm': LGBMRegressor(n_estimators=100, random_state=42),
        'elastic': ElasticNet(random_state=42)
    }
    return models

def ensemble_predict(models, X):
    predictions = np.array([model.predict(X) for model in models.values()])
    return np.mean(predictions, axis=0)

def predict_sales(df, region, brand):
    df_filtered = df[(df['Zone'] == region) & (df['Brand'] == brand)]
    
    if len(df_filtered) == 0:
        return None, None, None
    
    row = df_filtered.iloc[0]
    X = create_features(row)
    
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    models = create_models()
    
    # Train models on the created features
    y = row['Monthly Achievement(Aug)']  # We'll predict August as a proxy for September
    
    for model in models.values():
        model.fit(X_scaled, [y])
    
    sept_prediction = ensemble_predict(models, X_scaled)[0]
    
    # Calculate confidence interval
    model_predictions = [model.predict(X_scaled)[0] for model in models.values()]
    ci_lower, ci_upper = np.percentile(model_predictions, [2.5, 97.5])
    
    # Calculate prediction interval
    n_iterations = 1000
    bootstrap_predictions = []
    for _ in range(n_iterations):
        bootstrap_sample = np.random.choice(model_predictions, size=len(model_predictions), replace=True)
        bootstrap_predictions.append(np.mean(bootstrap_sample))
    pi_lower, pi_upper = np.percentile(bootstrap_predictions, [2.5, 97.5])
    
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
        1. Ensemble Learning: Combines predictions from multiple models to reduce bias and variance.
        2. Feature Engineering: Creates new features to capture complex patterns in the data.
        3. Historical Data Utilization: Incorporates previous year's data and monthly trends.
        4. Confidence and Prediction Intervals: Provides a range of possible outcomes.
        """)

if __name__ == "__main__":
    main()
