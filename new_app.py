import streamlit as st
import pandas as pd
import numpy as np
from sklearn.ensemble import RandomForestRegressor
from datetime import datetime
import calendar
import plotly.express as px
import plotly.graph_objects as go
from typing import Dict, Tuple, Optional, List
import os

# Set page config for a wider layout
st.set_page_config(
    page_title="Sales Forecasting Dashboard",
    page_icon="📈",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better aesthetics
st.markdown("""
    <style>
    .main {
        padding: 0rem 1rem;
    }
    .stSelectbox, .stNumberInput {
        margin-bottom: 1rem;
    }
    .reportview-container {
        background: #f0f2f6;
    }
    .sidebar .sidebar-content {
        background: #f9f9f9;
    }
    .st-emotion-cache-1y4p8pa {
        max-width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

class DataPreprocessor:
    @staticmethod
    def read_excel_skip_hidden(file) -> pd.DataFrame:
        """Read Excel file and skip hidden rows."""
        try:
            return pd.read_excel(file)
        except Exception as e:
            st.error(f"Error reading file: {str(e)}")
            return None

    @staticmethod
    def prepare_features(df: pd.DataFrame, current_month_data: Optional[Dict] = None) -> pd.DataFrame:
        """Prepare features for the prediction model."""
        features = pd.DataFrame()
        
        # Extract monthly sales
        months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
        for month in months:
            features[f'sales_{month}'] = df[f'Monthly Achievement({month})']
        
        # Previous year data
        features['prev_sep'] = df['Total Sep 2023']
        features['prev_oct'] = df['Total Oct 2023']
        features['prev_nov'] = df['Total Nov 2023']
        
        # Target achievement calculations
        for month in months:
            # Monthly targets
            features[f'month_target_{month}'] = df[f'Month Tgt ({month})']
            features[f'monthly_achievement_rate_{month}'] = (
                features[f'sales_{month}'] / features[f'month_target_{month}']
            )
            
            # AGS targets
            features[f'ags_target_{month}'] = df[f'AGS Tgt ({month})']
            features[f'ags_achievement_rate_{month}'] = (
                features[f'sales_{month}'] / features[f'ags_target_{month}']
            )
        
        # November targets
        features['month_target_nov'] = df['Month Tgt (Nov)']
        features['ags_target_nov'] = df['AGS Tgt (Nov)']
        
        # Calculate advanced metrics
        features = DataPreprocessor._calculate_advanced_metrics(features, current_month_data)
        
        return features

    @staticmethod
    def _calculate_advanced_metrics(features: pd.DataFrame, current_month_data: Optional[Dict]) -> pd.DataFrame:
        """Calculate advanced metrics for feature engineering."""
        # Achievement rates
        monthly_achievement_cols = [f'monthly_achievement_rate_{m}' for m in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']]
        ags_achievement_cols = [f'ags_achievement_rate_{m}' for m in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']]
        
        features['avg_monthly_achievement_rate'] = features[monthly_achievement_cols].mean(axis=1)
        features['avg_ags_achievement_rate'] = features[ags_achievement_cols].mean(axis=1)
        
        # Weighted achievement calculations
        weights = np.array([0.05, 0.1, 0.1, 0.15, 0.2, 0.2, 0.2])
        features['weighted_monthly_achievement_rate'] = np.average(
            features[monthly_achievement_cols], weights=weights, axis=1
        )
        features['weighted_ags_achievement_rate'] = np.average(
            features[ags_achievement_cols], weights=weights, axis=1
        )
        
        # Sales metrics
        sales_cols = [f'sales_{m}' for m in ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']]
        features['avg_monthly_sales'] = features[sales_cols].mean(axis=1)
        
        # Growth calculations
        months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep', 'Oct']
        for i in range(1, len(months)):
            features[f'growth_{months[i]}'] = features[f'sales_{months[i]}'] / features[f'sales_{months[i-1]}']
        
        # YoY calculations
        features['yoy_sep_growth'] = features['sales_Sep'] / features['prev_sep']
        features['yoy_oct_growth'] = features['sales_Oct'] / features['prev_oct']
        
        if current_month_data:
            DataPreprocessor._add_current_month_metrics(features, current_month_data)
        else:
            features['yoy_weighted_growth'] = (features['yoy_sep_growth'] * 0.4 + features['yoy_oct_growth'] * 0.6)
            
        features['target_achievement_rate'] = features['sales_Oct'] / features['month_target_Oct']
        return features

    @staticmethod
    def _add_current_month_metrics(features: pd.DataFrame, current_month_data: Dict):
        """Add current month specific metrics."""
        features['current_month_yoy_growth'] = current_month_data['current_year'] / current_month_data['previous_year']
        features['projected_full_month'] = (current_month_data['current_year'] / current_month_data['days_passed']) * current_month_data['total_days']
        features['current_month_daily_rate'] = current_month_data['current_year'] / current_month_data['days_passed']
        features['yoy_weighted_growth'] = (
            features['yoy_sep_growth'] * 0.3 + 
            features['yoy_oct_growth'] * 0.4 +
            features['current_month_yoy_growth'] * 0.3
        )

class SalesPredictor:
    def __init__(self):
        self.rf_model_monthly = RandomForestRegressor(n_estimators=100, random_state=42)
        self.rf_model_ags = RandomForestRegressor(n_estimators=100, random_state=42)

    def predict(self, 
               df: pd.DataFrame, 
               selected_zone: str, 
               selected_brand: str, 
               growth_weights: Dict[str, float], 
               method_weights: Dict[str, float], 
               current_month_data: Optional[Dict] = None) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """Generate sales predictions using multiple methods."""
        df_filtered = df[(df['Zone'] == selected_zone) & (df['Brand'] == selected_brand)]
        
        if len(df_filtered) == 0:
            st.error("No data available for the selected combination of Zone and Brand")
            return None, None
        
        features = DataPreprocessor.prepare_features(df_filtered, current_month_data)
        historical_data = self._prepare_historical_data(df_filtered, current_month_data)
        
        predictions = self._generate_predictions(
            features, growth_weights, method_weights, current_month_data
        )
        
        return predictions, historical_data

    def _generate_predictions(self, 
                            features: pd.DataFrame, 
                            growth_weights: Dict[str, float], 
                            method_weights: Dict[str, float],
                            current_month_data: Optional[Dict]) -> pd.DataFrame:
        """Generate predictions using multiple methods."""
        # RF Prediction
        rf_prediction = self._get_rf_prediction(features)
        
        # YoY Prediction
        yoy_prediction = features['prev_nov'] * features['yoy_weighted_growth']
        
        # Trend Prediction
        trend_prediction = self._calculate_trend_prediction(features, growth_weights, current_month_data)
        
        # Target-Based Prediction
        target_prediction = self._get_target_based_prediction(features, current_month_data)
        
        # Combine predictions
        final_prediction = (
            method_weights['rf'] * rf_prediction +
            method_weights['yoy'] * yoy_prediction +
            method_weights['trend'] * trend_prediction +
            method_weights['target'] * target_prediction
        )
        
        return pd.DataFrame({
            'RF_Prediction': rf_prediction,
            'YoY_Prediction': yoy_prediction,
            'Trend_Prediction': trend_prediction,
            'Target_Based_Prediction': target_prediction,
            'Final_Prediction': final_prediction
        })

    def _get_rf_prediction(self, features: pd.DataFrame) -> np.ndarray:
        """Get Random Forest prediction."""
        exclude_columns = [
            'month_target_nov', 'ags_target_nov',
            'avg_monthly_achievement_rate', 'avg_ags_achievement_rate',
            'weighted_monthly_achievement_rate', 'weighted_ags_achievement_rate',
            'avg_monthly_sales', 'yoy_sep_growth', 'yoy_oct_growth',
            'yoy_weighted_growth'
        ]
        feature_cols = [col for col in features.columns if col not in exclude_columns]
        
        self.rf_model_monthly.fit(
            features[feature_cols], 
            features['month_target_nov'] * features['weighted_monthly_achievement_rate']
        )
        self.rf_model_ags.fit(
            features[feature_cols], 
            features['ags_target_nov'] * features['weighted_ags_achievement_rate']
        )
        
        rf_prediction_monthly = self.rf_model_monthly.predict(features[feature_cols])
        rf_prediction_ags = self.rf_model_ags.predict(features[feature_cols])
        
        return (rf_prediction_monthly + rf_prediction_ags) / 2

    @staticmethod
    def _calculate_trend_prediction(features: pd.DataFrame, 
                                  growth_weights: Dict[str, float], 
                                  current_month_data: Optional[Dict]) -> np.ndarray:
        """Calculate trend-based prediction."""
        if current_month_data:
            adjusted_weights = {k: v * 0.7 for k, v in growth_weights.items()}
            adjusted_weights['current_month'] = 0.3
            
            base_weighted_growth = sum(
                features[month] * weight 
                for month, weight in adjusted_weights.items() 
                if month != 'current_month'
            ) / sum(adjusted_weights.values())
            
            current_month_growth = features['current_month_yoy_growth'].iloc[0]
            weighted_growth = (base_weighted_growth * 0.7 + current_month_growth * 0.3)
        else:
            weighted_growth = sum(
                features[month] * weight 
                for month, weight in growth_weights.items()
            ) / sum(growth_weights.values())
        
        return features['sales_Oct'] * weighted_growth

    @staticmethod
    def _get_target_based_prediction(features: pd.DataFrame, 
                                   current_month_data: Optional[Dict]) -> pd.Series:
        """Calculate target-based prediction."""
        if current_month_data:
            return (
                features['avg_monthly_sales'] * 
                features['target_achievement_rate'] * 
                (1 + (features['current_month_yoy_growth'] - 1) * 0.3)
            )
        return features['avg_monthly_sales'] * features['target_achievement_rate']

    @staticmethod
    def _prepare_historical_data(df_filtered: pd.DataFrame, 
                               current_month_data: Optional[Dict]) -> pd.DataFrame:
        """Prepare historical sales data."""
        if current_month_data:
            return pd.DataFrame({
                'Period': ['October 2023', 'November 2023', 'October 2024', 
                          f'November 2024 (First {current_month_data["days_passed"]} days)',
                          f'November 2023 (First {current_month_data["days_passed"]} days)'],
                'Sales': [
                    df_filtered['Total Oct 2023'].iloc[0],
                    df_filtered['Total Nov 2023'].iloc[0],
                    df_filtered['Monthly Achievement(Oct)'].iloc[0],
                    current_month_data['current_year'],
                    current_month_data['previous_year']
                ]
            })
        else:
            return pd.DataFrame({
                'Period': ['October 2023', 'November 2023', 'October 2024'],
                'Sales': [
                    df_filtered['Total Oct 2023'].iloc[0],
                    df_filtered['Total Nov 2023'].iloc[0],
                    df_filtered['Monthly Achievement(Oct)'].iloc[0]
                ]
            })

class Dashboard:
    def __init__(self):
        self.predictor = SalesPredictor()

    def run(self):
        """Run the Streamlit dashboard."""
        st.title("🚀 Advanced Sales Forecasting Dashboard")
        
        # Sidebar
        with st.sidebar:
            st.header("📊 Configuration")
            uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx'])
            
            if uploaded_file:
                self._process_uploaded_file(uploaded_file)

    def _process_uploaded_file(self, uploaded_file):
        """Process the uploaded file and display the dashboard."""
        df = DataPreprocessor.read_excel_skip_hidden(uploaded_file)
        
        if df is not None:
            # Sidebar controls
            with st.sidebar:
                st.subheader("🎯 Select Parameters")
                zone = st.selectbox("Zone", sorted(df['Zone'].unique()))
                brand = st.selectbox("Brand", sorted(df[df['Zone'] == zone]['Brand'].unique()))
                
                st.subheader("📅 Current Month Analysis")
                use_current_month = st.checkbox("Include current month data")
                
                current_month_data = self._get_current_month_data(use_current_month) if use_current_month else None
                
                st.subheader("⚖️ Configure Weights")
                growth_weights = self._get_growth_weights()
                method_weights = self._get_method_weights()
            
            # Main content
            col1, col2 = st.columns([2, 1])
            
            with col1:
                predictions, historical_data = self.predictor.predict(
                    df, zone, brand, growth_weights, method_weights, current_month_data
                )
                
                if predictions is not None and historical_data is not None:
                    self._display_predictions(predictions, historical_data, current_month_data)
            
                
            
            with col2:
                if predictions is not None:
                    self._display_metrics_and_charts(predictions, historical_data, current_month_data)

    def _get_current_month_data(self, use_current_month: bool) -> Optional[Dict]:
        """Get current month analysis data from user input."""
        if use_current_month:
            col1, col2 = st.columns(2)
            with col1:
                days_passed = st.number_input("Days passed", 1, 31, 1)
                current_year_sales = st.number_input("Current year sales", 0.0, None, 0.0)
            with col2:
                previous_year_sales = st.number_input("Previous year sales", 0.0, None, 0.0)
            
            if days_passed > 0 and current_year_sales > 0 and previous_year_sales > 0:
                return {
                    'days_passed': days_passed,
                    'total_days': calendar.monthrange(2024, 11)[1],
                    'current_year': current_year_sales,
                    'previous_year': previous_year_sales
                }
        return None

    def _get_growth_weights(self) -> Dict[str, float]:
        """Get growth weights from user input."""
        st.write("Growth Weights")
        col1, col2 = st.columns(2)
        
        weights = {}
        with col1:
            weights['growth_May'] = st.slider("May/Apr", 0.0, 1.0, 0.05, 0.05)
            weights['growth_June'] = st.slider("June/May", 0.0, 1.0, 0.10, 0.05)
            weights['growth_July'] = st.slider("July/June", 0.0, 1.0, 0.15, 0.05)
        
        with col2:
            weights['growth_Aug'] = st.slider("Aug/July", 0.0, 1.0, 0.20, 0.05)
            weights['growth_Sep'] = st.slider("Sep/Aug", 0.0, 1.0, 0.25, 0.05)
            weights['growth_Oct'] = st.slider("Oct/Sep", 0.0, 1.0, 0.25, 0.05)
        
        return weights

    def _get_method_weights(self) -> Dict[str, float]:
        """Get method weights from user input."""
        st.write("Method Weights")
        col1, col2 = st.columns(2)
        
        weights = {}
        with col1:
            weights['rf'] = st.slider("Random Forest", 0.0, 1.0, 0.4, 0.05)
            weights['yoy'] = st.slider("Year-over-Year", 0.0, 1.0, 0.1, 0.05)
        
        with col2:
            weights['trend'] = st.slider("Trend", 0.0, 1.0, 0.4, 0.05)
            weights['target'] = st.slider("Target-Based", 0.0, 1.0, 0.1, 0.05)
        
        return weights

    def _display_predictions(self, predictions: pd.DataFrame, historical_data: pd.DataFrame, current_month_data: Optional[Dict]):
        """Display predictions and historical data."""
        st.subheader("📊 Historical Sales Performance")
        historical_data['Growth'] = historical_data.apply(
            lambda x: f"{((x['Sales'] / historical_data.iloc[0]['Sales'] - 1) * 100):.1f}% vs Base" 
            if x.name > 0 else 'Base',
            axis=1
        )
        
        st.dataframe(
            historical_data.style.format({
                'Sales': '₹{:,.2f}'
            }),
            use_container_width=True
        )
        
        st.subheader("🎯 November 2024 Predictions")
        formatted_predictions = predictions.copy()
        for col in predictions.columns:
            formatted_predictions[col] = formatted_predictions[col].apply(lambda x: f"₹{x:,.2f}")
        
        st.dataframe(formatted_predictions, use_container_width=True)

    def _display_metrics_and_charts(self, predictions: pd.DataFrame, historical_data: pd.DataFrame, current_month_data: Optional[Dict]):
        """Display metrics and visualizations."""
        st.subheader("📈 Key Metrics")
        
        # Display key metrics in boxes
        col1, col2 = st.columns(2)
        with col1:
            st.metric(
                "Average Prediction",
                f"₹{predictions['Final_Prediction'].mean():,.2f}",
                delta=f"{((predictions['Final_Prediction'].mean() / historical_data['Sales'].iloc[1] - 1) * 100):.1f}% vs Last Year"
            )
        
        with col2:
            if current_month_data:
                current_performance = (current_month_data['current_year'] / current_month_data['previous_year'] - 1) * 100
                st.metric(
                    "Current Month Performance",
                    f"₹{current_month_data['current_year']:,.2f}",
                    f"{current_performance:.1f}% vs Last Year"
                )
        
        # Visualization of predictions
        st.subheader("📊 Prediction Comparison")
        fig = go.Figure()
        
        methods = ['RF_Prediction', 'YoY_Prediction', 'Trend_Prediction', 'Target_Based_Prediction', 'Final_Prediction']
        colors = ['rgb(99, 110, 250)', 'rgb(239, 85, 59)', 'rgb(0, 204, 150)', 'rgb(171, 99, 250)', 'rgb(255, 161, 90)']
        
        for method, color in zip(methods, colors):
            fig.add_trace(go.Bar(
                name=method.replace('_', ' '),
                y=[predictions[method].iloc[0]],
                marker_color=color
            ))
        
        fig.update_layout(
            title="Comparison of Different Prediction Methods",
            yaxis_title="Predicted Sales (₹)",
            showlegend=True,
            height=400,
            barmode='group'
        )
        
        st.plotly_chart(fig, use_container_width=True)
        
        if current_month_data:
            self._display_current_month_analysis(current_month_data, predictions)

    def _display_current_month_analysis(self, current_month_data: Dict, predictions: pd.DataFrame):
        """Display current month analysis."""
        st.subheader("📅 Current Month Analysis")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                "Days Completed",
                f"{current_month_data['days_passed']}/{current_month_data['total_days']}",
                f"{(current_month_data['days_passed']/current_month_data['total_days']*100):.1f}% Complete"
            )
        
        with col2:
            current_daily_rate = current_month_data['current_year'] / current_month_data['days_passed']
            previous_daily_rate = current_month_data['previous_year'] / current_month_data['days_passed']
            st.metric(
                "Current Daily Rate",
                f"₹{current_daily_rate:,.2f}",
                f"{((current_daily_rate/previous_daily_rate - 1) * 100):.1f}% vs Last Year"
            )
        
        with col3:
            projected_full_month = (current_month_data['current_year'] / current_month_data['days_passed']) * current_month_data['total_days']
            st.metric(
                "Projected Full Month",
                f"₹{projected_full_month:,.2f}",
                f"{((projected_full_month/predictions['Final_Prediction'].iloc[0] - 1) * 100):.1f}% vs Prediction"
            )

if __name__ == "__main__":
    dashboard = Dashboard()
    dashboard.run()
