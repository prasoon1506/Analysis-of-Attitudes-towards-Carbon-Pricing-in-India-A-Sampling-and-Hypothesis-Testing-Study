import streamlit as st
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import LabelEncoder
from sklearn.metrics import mean_absolute_percentage_error, mean_absolute_error
import warnings
import json
from itertools import product
from pathlib import Path
import plotly.graph_objects as go
import plotly.express as px

warnings.filterwarnings('ignore')

class WeightOptimizer:
    def __init__(self):
        self.optimal_weights = {}
        
    def generate_weight_combinations(self, num_weights, step=0.2):
        """Generate combinations of weights that sum to 1"""
        weights = np.arange(0, 1.01, step)
        combinations = []
        
        for w in product(weights, repeat=num_weights):
            if abs(sum(w) - 1.0) < 0.01:
                combinations.append(w)
                
        return combinations
    
    def calculate_prediction_error(self, actual, predicted):
        """Calculate error metrics"""
        mape = mean_absolute_percentage_error([actual], [predicted])
        mae = mean_absolute_error([actual], [predicted])
        return mape, mae
    
    def find_optimal_weights(self, df, zone, brand):
        """Find optimal weights for a specific zone and brand"""
        features = prepare_features_for_october(df)
        filtered_data = df[(df['Zone'] == zone) & (df['Brand'] == brand)]
        
        if len(filtered_data) == 0:
            return None
            
        actual_october = filtered_data['Monthly Achievement(Oct)'].iloc[0]
        
        # Generate weight combinations for 5 growth rates (May through September)
        growth_combinations = self.generate_weight_combinations(5, step=0.2)
        method_combinations = self.generate_weight_combinations(4, step=0.2)
        
        best_error = float('inf')
        best_weights = None
        
        for growth_weights in growth_combinations:
            growth_dict = {
                'growth_May': growth_weights[0],
                'growth_June': growth_weights[1],
                'growth_July': growth_weights[2],
                'growth_Aug': growth_weights[3],
                'growth_Sep': growth_weights[4],
                'growth_Oct': 1.0  # Oct/Sep growth will be used directly
            }
            
            for method_weights in method_combinations:
                method_dict = {
                    'rf': method_weights[0],
                    'yoy': method_weights[1],
                    'trend': method_weights[2],
                    'target': method_weights[3]
                }
                
                prediction = predict_october_sales(df, zone, brand, growth_dict, method_dict)
                
                if prediction is not None:
                    predicted_value = prediction['Final_Prediction'].iloc[0]
                    mape, mae = self.calculate_prediction_error(actual_october, predicted_value)
                    
                    if mape < best_error:
                        best_error = mape
                        best_weights = {
                            'growth_weights': growth_dict,
                            'method_weights': method_dict,
                            'mape': mape,
                            'mae': mae
                        }
        
        return best_weights
    
    def optimize_all_combinations(self, df):
        """Find optimal weights for all zone-brand combinations"""
        for zone in df['Zone'].unique():
            for brand in df[df['Zone'] == zone]['Brand'].unique():
                best_weights = self.find_optimal_weights(df, zone, brand)
                if best_weights is not None:
                    self.optimal_weights[f"{zone}_{brand}"] = best_weights
        
        self.save_weights()
    
    def save_weights(self):
        """Save optimal weights to a JSON file"""
        with open('optimal_weights.json', 'w') as f:
            json.dump(self.optimal_weights, f)
    
    def load_weights(self):
        """Load optimal weights from JSON file"""
        try:
            with open('optimal_weights.json', 'r') as f:
                self.optimal_weights = json.load(f)
        except FileNotFoundError:
            self.optimal_weights = {}

def prepare_features_for_prediction(df, include_october=True):
    """Prepare features for prediction"""
    features = pd.DataFrame()
    
    # Extract monthly sales
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']
    if include_october:
        months.append('Oct')
        
    for month in months:
        features[f'sales_{month}'] = df[f'Monthly Achievement({month})']
    
    # Add previous year data
    features['prev_sep'] = df['Total Sep 2023']
    features['prev_oct'] = df['Total Oct 2023']
    features['prev_nov'] = df['Total Nov 2023']
    
    # Add target information
    for month in months:
        features[f'month_target_{month}'] = df[f'Month Tgt ({month})']
        features[f'ags_target_{month}'] = df[f'AGS Tgt ({month})']
    
    # Calculate additional features
    features['avg_monthly_sales'] = features[[f'sales_{m}' for m in months]].mean(axis=1)
    
    # Calculate month-over-month growth rates
    for i in range(1, len(months)):
        features[f'growth_{months[i]}'] = features[f'sales_{months[i]}'] / features[f'sales_{months[i-1]}']
    
    # Calculate YoY growth rates
    if include_october:
        features['yoy_oct_growth'] = features['sales_Oct'] / features['prev_oct']
    features['yoy_sep_growth'] = features['sales_Sep'] / features['prev_sep']
    
    # Target achievement rates
    if include_october:
        features['target_achievement_rate'] = features['sales_Oct'] / features['month_target_Oct']
    else:
        features['target_achievement_rate'] = features['sales_Sep'] / features['month_target_Sep']
    
    return features

def predict_november_sales(df, zone, brand, growth_weights, method_weights):
    """Predict November sales using multiple methods and weighted averaging"""
    # Filter data for selected zone and brand
    data = df[(df['Zone'] == zone) & (df['Brand'] == brand)].copy()
    
    if len(data) == 0:
        return None
    
    features = prepare_features_for_prediction(data, include_october=True)
    
    # Calculate weighted growth rate
    weighted_growth = sum(
        growth_weights[f'growth_{month}'] * features[f'growth_{month}'].iloc[0]
        for month in ['May', 'June', 'July', 'Aug', 'Sep', 'Oct']
    )
    
    # Method 1: Random Forest prediction
    rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
    X = features.drop(['yoy_oct_growth', 'target_achievement_rate'], axis=1)
    y = data['Monthly Achievement(Oct)']
    rf_model.fit(X, y)
    rf_prediction = rf_model.predict(X)[0] * weighted_growth
    
    # Method 2: Year-over-Year growth
    yoy_prediction = data['Total Nov 2023'].iloc[0] * features['yoy_oct_growth'].iloc[0]
    
    # Method 3: Trend-based prediction
    trend_prediction = data['Monthly Achievement(Oct)'].iloc[0] * weighted_growth
    
    # Method 4: Target-based prediction
    target_prediction = data['Month Tgt (Nov)'].iloc[0] * features['target_achievement_rate'].iloc[0]
    
    # Combine predictions using method weights
    final_prediction = (
        method_weights['rf'] * rf_prediction +
        method_weights['yoy'] * yoy_prediction +
        method_weights['trend'] * trend_prediction +
        method_weights['target'] * target_prediction
    )
    
    # Create results DataFrame
    results = pd.DataFrame({
        'Method': ['Random Forest', 'Year-over-Year', 'Trend-based', 'Target-based', 'Final Prediction'],
        'Prediction': [rf_prediction, yoy_prediction, trend_prediction, target_prediction, final_prediction]
    })
    
    return results

def create_sales_history_plot(df, zone, brand):
    """Create a line plot of historical sales data"""
    data = df[(df['Zone'] == zone) & (df['Brand'] == brand)].copy()
    
    if len(data) == 0:
        return None
        
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
    sales_data = {month: data[f'Monthly Achievement({month})'].iloc[0] for month in months}
    
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=list(sales_data.keys()),
        y=list(sales_data.values()),
        mode='lines+markers',
        name='Actual Sales'
    ))
    
    fig.update_layout(
        title=f'Historical Sales Data for {brand} in {zone}',
        xaxis_title='Month',
        yaxis_title='Sales Amount',
        template='plotly_white'
    )
    
    return fig

def main():
    st.set_page_config(page_title="Sales Forecasting App", layout="wide")
    
    st.title("Sales Forecasting Model")
    st.markdown("""
    This app predicts November sales based on historical data and various prediction methods.
    Upload your Excel file and configure the weights to generate predictions.
    """)
    
    # Initialize weight optimizer
    weight_optimizer = WeightOptimizer()
    weight_optimizer.load_weights()
    
    # File upload
    uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        # Sidebar for controls
        st.sidebar.header("Controls")
        
        # Calibration section
        st.sidebar.subheader("Weight Calibration")
        if st.sidebar.button("Calibrate Weights"):
            with st.spinner("Calibrating weights for all zone-brand combinations..."):
                weight_optimizer.optimize_all_combinations(df)
            st.success("Calibration complete!")
        
        # Zone and Brand selection
        zones = sorted(df['Zone'].unique())
        brands = sorted(df['Brand'].unique())
        
        selected_zone = st.sidebar.selectbox("Select Zone", zones)
        selected_brand = st.sidebar.selectbox("Select Brand", brands)
        
        # Option to use recommended weights
        use_recommended = st.sidebar.checkbox("Use Recommended Weights")
        
        # Weight configuration
        st.sidebar.subheader("Configure Weights")
        
        # Initialize weights
        growth_weights = {}
        method_weights = {}
        
        # Load recommended weights if available and selected
        if use_recommended:
            key = f"{selected_zone}_{selected_brand}"
            if key in weight_optimizer.optimal_weights:
                recommended = weight_optimizer.optimal_weights[key]
                growth_weights = recommended['growth_weights']
                method_weights = recommended['method_weights']
            else:
                st.warning("No recommended weights available for this combination.")
                use_recommended = False
        
        # Growth weight sliders
        st.sidebar.subheader("Growth Weights")
        if not use_recommended:
            growth_weights = {
                'growth_May': st.sidebar.slider("May/Apr Weight", 0.0, 1.0, 0.1, 0.05),
                'growth_June': st.sidebar.slider("June/May Weight", 0.0, 1.0, 0.15, 0.05),
                'growth_July': st.sidebar.slider("July/June Weight", 0.0, 1.0, 0.2, 0.05),
                'growth_Aug': st.sidebar.slider("Aug/July Weight", 0.0, 1.0, 0.25, 0.05),
                'growth_Sep': st.sidebar.slider("Sep/Aug Weight", 0.0, 1.0, 0.3, 0.05),
                'growth_Oct': 1.0  # Fixed weight for Oct/Sep
            }
        
        # Method weight sliders
        st.sidebar.subheader("Method Weights")
        if not use_recommended:
            method_weights = {
                'rf': st.sidebar.slider("Random Forest Weight", 0.0, 1.0, 0.4, 0.05),
                'yoy': st.sidebar.slider("Year-over-Year Weight", 0.0, 1.0, 0.1, 0.05),
                'trend': st.sidebar.slider("Trend Weight", 0.0, 1.0, 0.4, 0.05),
                'target': st.sidebar.slider("Target Weight", 0.0, 1.0, 0.1, 0.05)
            }
        
        # Validate weights
        growth_sum = sum(list(growth_weights.values())[:-1])  # Exclude Oct/Sep weight
        method_sum = sum(method_weights.values())
        
        # Main content area
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Historical Sales Data")
            fig = create_sales_history_plot(df, selected_zone, selected_brand)
            if fig is not None:
                st.plotly_chart(fig)
        
        with col2:
            st.subheader("November Sales Prediction")
            
            if abs(growth_sum - 1.0) > 0.01:
                st.error("Growth weights (May through September) must sum to 1.0")
            elif abs(method_sum - 1.0) > 0.01:
                st.error("Method weights must sum to 1.0")
            else:
                predictions = predict_november_sales(
                    df, selected_zone, selected_brand,
                    growth_weights, method_weights
                )
                
                if predictions is not None:
                    # Format predictions
                    predictions['Prediction'] = predictions['Prediction'].round(2)
                    predictions['Prediction'] = predictions['Prediction'].apply(lambda x: f"₹{x:,.2f}")
                    
                    # Display predictions
                    st.dataframe(predictions.set_index('Method'))
                    
                    # Display historical error metrics if using recommended weights
                    if use_recommended:
                        key = f"{selected_zone}_{selected_brand}"
                        if key in weight_optimizer.optimal_weights:
                            st.subheader("Historical Error Metrics (October Validation)")
                            metrics = weight_optimizer.optimal_weights[key]
                            col1, col2 = st.columns(2)
                            col1.metric("MAPE", f"{metrics['mape']:.2%}")
                            col2.metric("MAE", f"₹{metrics['mae']:,.2f}")

if __name__ == "__main__":
    main()
