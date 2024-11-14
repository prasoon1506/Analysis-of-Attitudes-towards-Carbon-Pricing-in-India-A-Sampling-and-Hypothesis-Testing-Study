# [Previous imports and functions remain the same until the main() function]
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
    
                            # Display detailed predictions table
                            st.subheader("Detailed Predictions")
                            
                            # Create a styled dataframe with proper formatting
                            styled_predictions = predictions.copy()
                            
                            # Format only numeric columns
                            numeric_cols = predictions.select_dtypes(include=['float64', 'int64']).columns
                            for col in numeric_cols:
                                styled_predictions[col] = styled_predictions[col].map('{:,.2f}'.format)
                            
                            st.dataframe(styled_predictions)
                            
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
