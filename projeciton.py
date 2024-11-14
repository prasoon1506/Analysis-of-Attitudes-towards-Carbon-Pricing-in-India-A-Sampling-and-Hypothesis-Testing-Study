import streamlit as st
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import mean_absolute_percentage_error
import warnings
from openpyxl import load_workbook
import plotly.express as px
import plotly.graph_objects as go

warnings.filterwarnings('ignore')

def read_excel_skip_hidden(uploaded_file):
    """
    Read Excel file while skipping hidden rows
    """
    # Save uploaded file to a temporary file
    wb = load_workbook(uploaded_file)
    ws = wb.active
    hidden_rows = [i + 1 for i in range(ws.max_row) if ws.row_dimensions[i + 1].hidden]
    df = pd.read_excel(
        uploaded_file,
        skiprows=hidden_rows
    )
    return df

def prepare_features(df):
    """
    Prepare features for the prediction model
    """
    features = pd.DataFrame()
    
    # Extract current year monthly sales (Apr to Oct)
    for month in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']:
        features[f'sales_{month}'] = df[f'Monthly Achievement({month})']
    
    # Add previous year September, October and November sales
    features['prev_sep'] = df['Total Sep 2023']
    features['prev_oct'] = df['Total Oct 2023']
    features['prev_nov'] = df['Total Nov 2023']
    
    # Add target information
    for month in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']:
        features[f'month_target_{month}'] = df[f'Month Tgt ({month})']
        features[f'ags_target_{month}'] = df[f'AGS Tgt ({month})']
    
    # Calculate additional features
    features['avg_monthly_sales'] = features[[f'sales_{m}' for m in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']]].mean(axis=1)
    
    # Calculate month-over-month growth rates
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
    for i in range(1, len(months)):
        features[f'growth_{months[i]}'] = features[f'sales_{months[i]}'] / features[f'sales_{months[i-1]}']
    
    # Calculate YoY growth rates
    features['yoy_sep_growth'] = features['sales_Sep'] / features['prev_sep']
    features['yoy_oct_growth'] = features['sales_Oct'] / features['prev_oct']
    features['yoy_weighted_growth'] = (features['yoy_sep_growth'] * 0.4 + features['yoy_oct_growth'] * 0.6)
    features['target_achievement_rate'] = features['sales_Oct'] / features['month_target_Oct']
    
    return features

def calculate_trend_prediction(features, growth_weights):
    """
    Calculate trend-based prediction using weighted average of historical growth rates
    """
    weighted_growth = sum(features[month] * weight 
                        for month, weight in growth_weights.items()) / sum(growth_weights.values())
    return features['sales_Oct'] * weighted_growth

def predict_november_sales(df, selected_zone, selected_brand, growth_weights, method_weights):
    """
    Generate November sales predictions using multiple methods
    """
    # Filter data for selected zone and brand
    df_filtered = df[(df['Zone'] == selected_zone) & (df['Brand'] == selected_brand)]
    
    if len(df_filtered) == 0:
        st.error("No data available for the selected combination of Zone and Brand")
        return None
        
    features = prepare_features(df_filtered)
    
    # Method 1: Random Forest
    rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
    feature_cols = [col for col in features.columns if col not in 
                   ['avg_monthly_sales', 'yoy_sep_growth', 'yoy_oct_growth', 'yoy_weighted_growth', 
                    'target_achievement_rate'] and not col.startswith('growth_')]
    
    rf_model.fit(features[feature_cols], features['sales_Oct'])
    rf_prediction = rf_model.predict(features[feature_cols])
    
    # Method 2: Enhanced Year-over-Year Growth
    yoy_prediction = features['prev_nov'] * features['yoy_weighted_growth']
    
    # Method 3: Enhanced Trend Based
    trend_prediction = calculate_trend_prediction(features, growth_weights)
    
    # Method 4: Target-Based Prediction
    target_based_prediction = features['avg_monthly_sales'] * features['target_achievement_rate']
    
    # Combine predictions
    final_prediction = (
        method_weights['rf'] * rf_prediction +
        method_weights['yoy'] * yoy_prediction +
        method_weights['trend'] * trend_prediction +
        method_weights['target'] * target_based_prediction
    )
    
    return pd.DataFrame({
        'Zone': df_filtered['Zone'],
        'Brand': df_filtered['Brand'],
        'RF_Prediction': rf_prediction,
        'YoY_Prediction': yoy_prediction,
        'Trend_Prediction': trend_prediction,
        'Target_Based_Prediction': target_based_prediction,
        'Final_Prediction': final_prediction
    })

def create_prediction_charts(predictions):
    """
    Create visualizations for the predictions
    """
    # Method comparison chart
    methods_df = predictions.melt(
        id_vars=['Zone', 'Brand'],
        value_vars=['RF_Prediction', 'YoY_Prediction', 'Trend_Prediction', 'Target_Based_Prediction'],
        var_name='Method',
        value_name='Prediction'
    )
    
    fig_methods = px.bar(
        methods_df,
        x='Method',
        y='Prediction',
        title='Prediction by Method',
        template='plotly_white'
    )
    fig_methods.update_layout(yaxis_title='Predicted Sales (₹)')
    
    # Final prediction gauge
    fig_gauge = go.Figure(go.Indicator(
        mode="gauge+number",
        value=predictions['Final_Prediction'].mean(),
        title={'text': "Final Prediction (₹)"},
        gauge={'axis': {'range': [None, predictions['Final_Prediction'].mean() * 1.5]},
               'bar': {'color': "darkblue"}}
    ))
    
    return fig_methods, fig_gauge

def main():
    st.set_page_config(page_title="Sales Forecasting Model", layout="wide")
    
    # Add custom CSS
    st.markdown("""
        <style>
        .stApp {
            max-width: 1200px;
            margin: 0 auto;
        }
        .stButton button {
            width: 100%;
        }
        .stAlert {
            padding: 1rem;
            margin: 1rem 0;
        }
        </style>
    """, unsafe_allow_html=True)
    
    st.title("📈 Sales Forecasting Model")
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            df = read_excel_skip_hidden(uploaded_file)
            
            # Create two columns for zone and brand selection
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("Zone Selection")
                zones = sorted(df['Zone'].unique())
                selected_zone = st.selectbox('Select Zone', zones)
            
            with col2:
                st.subheader("Brand Selection")
                brands = sorted(df[df['Zone'] == selected_zone]['Brand'].unique())
                selected_brand = st.selectbox('Select Brand', brands)
            
            # Create tabs for weights configuration
            tab1, tab2 = st.tabs(["Growth Weights", "Method Weights"])
            
            with tab1:
                st.subheader("Configure Growth Weights")
                growth_weights = {}
                default_growth_weights = {
                    'growth_May': 0.05, 'growth_June': 0.1, 'growth_July': 0.15,
                    'growth_Aug': 0.2, 'growth_Sep': 0.25, 'growth_Oct': 0.25
                }
                
                for month, default_weight in default_growth_weights.items():
                    growth_weights[month] = st.slider(
                        f"{month.replace('growth_', '')} Growth Weight",
                        0.0, 1.0, default_weight, 0.05
                    )
                
                if abs(sum(growth_weights.values()) - 1.0) > 0.01:
                    st.warning("⚠️ Growth weights should sum to 1")
            
            with tab2:
                st.subheader("Configure Method Weights")
                method_weights = {}
                default_method_weights = {
                    'rf': 0.4, 'yoy': 0.1, 'trend': 0.4, 'target': 0.1
                }
                
                for method, default_weight in default_method_weights.items():
                    method_weights[method] = st.slider(
                        f"{method.upper()} Weight",
                        0.0, 1.0, default_weight, 0.05
                    )
                
                if abs(sum(method_weights.values()) - 1.0) > 0.01:
                    st.warning("⚠️ Method weights should sum to 1")
            
            # Generate predictions button
            if st.button("Generate Predictions", type="primary"):
                if abs(sum(growth_weights.values()) - 1.0) > 0.01 or abs(sum(method_weights.values()) - 1.0) > 0.01:
                    st.error("Please adjust weights to sum to 1 before generating predictions")
                else:
                    with st.spinner("Generating predictions..."):
                        predictions = predict_november_sales(
                            df, selected_zone, selected_brand,
                            growth_weights, method_weights
                        )
                        
                        if predictions is not None:
                            # Create visualization columns
                            chart_col1, chart_col2 = st.columns(2)
                            
                            # Generate and display charts
                            fig_methods, fig_gauge = create_prediction_charts(predictions)
                            
                            with chart_col1:
                                st.plotly_chart(fig_methods, use_container_width=True)
                            
                            with chart_col2:
                                st.plotly_chart(fig_gauge, use_container_width=True)
                            
                            # Display metrics
                            st.subheader("Summary Metrics")
                            metric_col1, metric_col2, metric_col3 = st.columns(3)
                            
                            with metric_col1:
                                st.metric("Average Prediction", f"₹{predictions['Final_Prediction'].mean():,.2f}")
                            
                            with metric_col2:
                                st.metric("Minimum Prediction", f"₹{predictions['Final_Prediction'].min():,.2f}")
                            
                            with metric_col3:
                                st.metric("Maximum Prediction", f"₹{predictions['Final_Prediction'].max():,.2f}")
                            
                            # Display detailed predictions table
                            st.subheader("Detailed Predictions")
                            st.dataframe(predictions.style.format("{:,.2f}"))
                            
                            # Add download button for predictions
                            csv = predictions.to_csv(index=False)
                            st.download_button(
                                label="Download Predictions as CSV",
                                data=csv,
                                file_name="sales_predictions.csv",
                                mime="text/csv"
                            )
        
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.info("Please make sure you've uploaded a valid Excel file with the correct format")

if __name__ == "__main__":
    main()
