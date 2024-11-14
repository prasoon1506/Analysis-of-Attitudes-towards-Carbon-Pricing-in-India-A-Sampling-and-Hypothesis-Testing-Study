import streamlit as st
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_percentage_error
import json
from pathlib import Path
import plotly.express as px
import plotly.graph_objects as go
import warnings

warnings.filterwarnings('ignore')

class WeightOptimizer:
    def __init__(self):
        self.optimal_weights = {}
        
    def generate_weight_combinations(self, num_weights, step=0.2):
        """Generate combinations of weights that sum to 1"""
        weights = np.arange(0, 1.01, step)
        combinations = []
        for combo in product(weights, repeat=num_weights):
            if abs(sum(combo) - 1.0) < 0.01:
                combinations.append(combo)
        return combinations

    def prepare_features_for_october(self, df):
        """Prepare features for October prediction"""
        features = pd.DataFrame()
        
        # Extract monthly sales up to September
        for month in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']:
            features[f'sales_{month}'] = df[f'Monthly Achievement({month})']
        
        features['prev_sep'] = df['Total Sep 2023']
        features['prev_oct'] = df['Total Oct 2023']
        
        # Add target information
        for month in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']:
            features[f'month_target_{month}'] = df[f'Month Tgt ({month})']
            features[f'ags_target_{month}'] = df[f'AGS Tgt ({month})']
        
        features['avg_monthly_sales'] = features[[f'sales_{m}' for m in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']]].mean(axis=1)
        
        # Calculate month-over-month growth rates
        months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']
        for i in range(1, len(months)):
            features[f'growth_{months[i]}'] = features[f'sales_{months[i]}'] / features[f'sales_{months[i-1]}']
        features['yoy_sep_growth'] = features['sales_Sep'] / features['prev_sep']
        features['yoy_weighted_growth'] = + features['yoy_sep_growth'] * 1
        
        features['target_achievement_rate'] = features['sales_Sep'] / features['month_target_Sep']
        features['actual_oct'] = df['Monthly Achievement(Oct)']
        
        return features

    def predict_october(self, features, growth_weights, method_weights):
        """Generate October prediction using the given weights"""
        # Random Forest prediction
        rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
        feature_cols = [col for col in features.columns if col not in 
                       ['avg_monthly_sales', 'yoy_aug_growth', 'yoy_sep_growth', 'yoy_weighted_growth',
                        'target_achievement_rate', 'actual_oct'] and not col.startswith('growth_')]
        
        rf_model.fit(features[feature_cols], features['sales_Sep'])
        rf_prediction = rf_model.predict(features[feature_cols])
        
        # YoY prediction
        yoy_prediction = features['prev_oct'] * features['yoy_weighted_growth']
        
        # Trend prediction
        weighted_growth = sum(features[month] * weight 
                            for month, weight in growth_weights.items()) / sum(growth_weights.values())
        trend_prediction = features['sales_Sep'] * weighted_growth
        
        # Target-based prediction
        target_based_prediction = features['avg_monthly_sales'] * features['target_achievement_rate']
        
        # Combine predictions
        final_prediction = (
            method_weights['rf'] * rf_prediction +
            method_weights['yoy'] * yoy_prediction +
            method_weights['trend'] * trend_prediction +
            method_weights['target'] * target_based_prediction
        )
        
        return final_prediction

    def find_optimal_weights(self, df, zone, brand):
        """Find optimal weights for a given zone and brand"""
        df_filtered = df[(df['Zone'] == zone) & (df['Brand'] == brand)]
        if len(df_filtered) == 0:
            return None
            
        features = self.prepare_features_for_october(df_filtered)
        growth_weight_combos = self.generate_weight_combinations(5, step=0.2)
        method_weight_combos = self.generate_weight_combinations(4, step=0.2)
        
        best_mape = float('inf')
        best_weights = None
        
        for growth_combo in growth_weight_combos:
            growth_weights = {
                'growth_May': growth_combo[0],
                'growth_June': growth_combo[1],
                'growth_July': growth_combo[2],
                'growth_Aug': growth_combo[3],
                'growth_Sep': growth_combo[4]
            }
            
            for method_combo in method_weight_combos:
                method_weights = {
                    'rf': method_combo[0],
                    'yoy': method_combo[1],
                    'trend': method_combo[2],
                    'target': method_combo[3]
                }
                
                prediction = self.predict_october(features, growth_weights, method_weights)
                mape = mean_absolute_percentage_error(features['actual_oct'], prediction)
                
                if mape < best_mape:
                    best_mape = mape
                    best_weights = {
                        'growth_weights': growth_weights,
                        'method_weights': method_weights,
                        'mape': mape
                    }
        
        return best_weights

def main():
    st.set_page_config(page_title="Sales Forecasting Model", layout="wide")
    
    # Add custom CSS
    st.markdown("""
        <style>
        .stSlider p {
            font-size: 1.1rem;
            color: #4A4A4A;
        }
        .block-container {
            padding-top: 2rem;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.title("🎯 Sales Forecasting Model")
    st.markdown("---")
    
    # Initialize session state
    if 'optimizer' not in st.session_state:
        st.session_state.optimizer = WeightOptimizer()
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            
            # Sidebar for controls
            with st.sidebar:
                st.header("Control Panel")
                
                # Zone and Brand selection
                zones = sorted(df['Zone'].unique())
                selected_zone = st.selectbox("Select Zone", zones)
                
                brands = sorted(df[df['Zone'] == selected_zone]['Brand'].unique())
                selected_brand = st.selectbox("Select Brand", brands)
                
                # Weight optimization
                st.subheader("Weight Optimization")
                use_recommended = st.checkbox("Use Recommended Weights")
                
                # Initialize default weights
                default_growth_weights = {
                    'growth_May': 0.05,
                    'growth_June': 0.1,
                    'growth_July': 0.15,
                    'growth_Aug': 0.2,
                    'growth_Sep': 0.25,
                    'growth_Oct': 0.25
                }
                
                default_method_weights = {
                    'rf': 0.4,
                    'yoy': 0.1,
                    'trend': 0.4,
                    'target': 0.1
                }
                
                # Update weights if recommended
                if use_recommended:
                    key = f"{selected_zone}_{selected_brand}"
                    if key in st.session_state.optimizer.optimal_weights:
                        weights = st.session_state.optimizer.optimal_weights[key]
                        default_growth_weights.update(weights['growth_weights'])
                        default_method_weights.update(weights['method_weights'])
                        st.success(f"Optimal MAPE: {weights['mape']:.2%}")
            
            # Main content area
            col1, col2 = st.columns([1, 1])
            
            with col1:
                st.subheader("Growth Weights")
                growth_weights = {}
                for month, default_value in default_growth_weights.items():
                    growth_weights[month] = st.slider(
                        f"{month.replace('growth_', '')} Growth Weight",
                        0.0, 1.0, default_value, 0.05,
                        key=f"growth_{month}"
                    )
                
                # Validate growth weights
                total_growth = sum(growth_weights.values())
                if abs(total_growth - 1.0) > 0.01:
                    st.warning(f"Growth weights sum to {total_growth:.2f}. Please adjust to sum to 1.0")
            
            with col2:
                st.subheader("Method Weights")
                method_weights = {}
                method_names = {
                    'rf': 'Random Forest',
                    'yoy': 'Year over Year',
                    'trend': 'Trend Based',
                    'target': 'Target Based'
                }
                
                for method, default_value in default_method_weights.items():
                    method_weights[method] = st.slider(
                        f"{method_names[method]} Weight",
                        0.0, 1.0, default_value, 0.05,
                        key=f"method_{method}"
                    )
                
                # Validate method weights
                total_method = sum(method_weights.values())
                if abs(total_method - 1.0) > 0.01:
                    st.warning(f"Method weights sum to {total_method:.2f}. Please adjust to sum to 1.0")
            
            # Generate predictions when weights are valid
            if abs(total_growth - 1.0) <= 0.01 and abs(total_method - 1.0) <= 0.01:
                if st.button("Generate Predictions", type="primary"):
                    # Generate predictions logic here
                    features = st.session_state.optimizer.prepare_features_for_october(
                        df[(df['Zone'] == selected_zone) & (df['Brand'] == selected_brand)]
                    )
                    
                    predictions = st.session_state.optimizer.predict_october(
                        features, growth_weights, method_weights
                    )
                    
                    # Display results
                    st.markdown("---")
                    st.subheader("Prediction Results")
                    
                    # Summary metrics
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Average Prediction", f"₹{predictions.mean():,.2f}")
                    with col2:
                        st.metric("Minimum Prediction", f"₹{predictions.min():,.2f}")
                    with col3:
                        st.metric("Maximum Prediction", f"₹{predictions.max():,.2f}")
                    
                    # Visualization
                    st.subheader("Prediction Visualization")
                    fig = go.Figure()
                    
                    fig.add_trace(go.Scatter(
                        y=features['actual_oct'],
                        name="Actual October Sales",
                        mode="markers",
                        marker=dict(size=10, color="blue")
                    ))
                    
                    fig.add_trace(go.Scatter(
                        y=predictions,
                        name="Predicted Sales",
                        mode="markers",
                        marker=dict(size=10, color="red")
                    ))
                    
                    fig.update_layout(
                        title="Actual vs Predicted Sales",
                        xaxis_title="Store Index",
                        yaxis_title="Sales (₹)",
                        showlegend=True
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display detailed predictions
                    comparison_df = pd.DataFrame({
                        'Actual_October': features['actual_oct'],
                        'Predicted_October': predictions,
                        'Difference': features['actual_oct'] - predictions,
                        'Percentage_Difference': ((features['actual_oct'] - predictions) / features['actual_oct']) * 100
                    })
                    
                    st.dataframe(
                        comparison_df.style.format({
                            'Actual_October': '₹{:,.2f}',
                            'Predicted_October': '₹{:,.2f}',
                            'Difference': '₹{:,.2f}',
                            'Percentage_Difference': '{:,.2f}%'
                        })
                    )
                    
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
    
    else:
        # Display instructions when no file is uploaded
        st.info("Please upload an Excel file to begin the analysis.")
        st.markdown("""
        ### Instructions:
        1. Upload your Excel file using the button above
        2. Select your Zone and Brand from the sidebar
        3. Adjust the weights or use recommended weights
        4. Click 'Generate Predictions' to see the results
        """)

if __name__ == "__main__":
    main()
