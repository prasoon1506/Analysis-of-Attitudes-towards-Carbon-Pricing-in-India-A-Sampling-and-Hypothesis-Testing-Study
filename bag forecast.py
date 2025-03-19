import streamlit as st
import pandas as pd
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.seasonal import seasonal_decompose
from sklearn.ensemble import RandomForestRegressor
import plotly.graph_objects as go
from datetime import datetime
import warnings
warnings.filterwarnings('ignore')
class BagDemandForecaster:
    def __init__(self, historical_data):
        self.data = historical_data.copy()
        self.data['Year'] = self.data['Date'].dt.year
        self.data['Month'] = self.data['Date'].dt.month
    def _calculate_seasonal_indices(self):
        try:
            decomposition = seasonal_decompose(self.data['Usage'], period=12, model='multiplicative')
            return pd.Series(decomposition.seasonal).mean()
        except:
            return 1.0
    def _extrapolate_february(self, mar_partial_data):
        days_in_mar = 31
        current_days = 18
        daily_rate = mar_partial_data / current_days
        return daily_rate * days_in_mar
    def _get_recent_trend(self, window=6):
        recent_data = self.data.tail(window)
        return (recent_data['Usage'].iloc[-1] - recent_data['Usage'].iloc[0]) / window
    def generate_forecast(self, feb_partial_data):
        mar_full = self._extrapolate_february(mar_partial_data)
        seasonal_indices = self._calculate_seasonal_indices()
        april_seasonal_factor = 1.0 if isinstance(seasonal_indices, float) else seasonal_indices.get(4, 1.0)
        try:
            model_hw = ExponentialSmoothing(self.data['Usage'],seasonal_periods=12,trend='add',seasonal='mul').fit()
            hw_forecast = model_hw.forecast(1)[0]
        except:
            hw_forecast = self.data['Usage'].mean()
        try:
            rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
            X = pd.DataFrame({'Month': self.data['Month'],'Year': self.data['Year'],'Lag1': self.data['Usage'].shift(1),'Lag12': self.data['Usage'].shift(12),'Trend': range(len(self.data))}).dropna()
            y = self.data['Usage'].iloc[12:]
            rf_model.fit(X, y)
            april_features = pd.DataFrame({'Month': [4],'Year': [2025],'Lag1': [feb_full],'Lag12': [self.data['Usage'].iloc[-12]],'Trend': [len(self.data)]})
            rf_forecast = rf_model.predict(march_features)[0]
        except:
            rf_forecast = self.data['Usage'].mean() * april_seasonal_factor
        trend = self._get_recent_trend()
        trend_forecast = mar_full + trend
        weights = {'holt_winters': 0.4,'random_forest': 0.4,'trend_based': 0.2}
        combined_forecast = (weights['holt_winters'] * hw_forecast +weights['random_forest'] * rf_forecast +weights['trend_based'] * trend_forecast)
        forecasts = np.array([hw_forecast, rf_forecast, trend_forecast])
        std_dev = np.std(forecasts)
        z_score = 1.96
        return {'point_forecast': combined_forecast,'lower_bound': combined_forecast - z_score * std_dev,'upper_bound': combined_forecast + z_score * std_dev,'individual_forecasts': {'holt_winters': hw_forecast,'random_forest': rf_forecast,'trend_based': trend_forecast},'march_extrapolated': mar_full,'seasonal_factor': april_seasonal_factor}
def main():
    st.set_page_config(page_title="Bag Demand Forecasting Test", layout="wide")
    st.title("🔮 Bag Demand Forecasting Test Application")
    st.write("Upload your data file and select a plant and bag to see the forecast.")
    uploaded_file = st.file_uploader("Upload Excel file", type=['xlsx', 'xls'])
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            df = df.iloc[:, 1:]
            plants = sorted(df['Cement Plant Sname'].unique())
            selected_plant = st.selectbox('Select Plant:', plants)
            plant_bags = sorted(df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique())
            selected_bag = st.selectbox('Select Bag:', plant_bags)
            if st.button('Generate Forecast'):
                st.markdown("---")
                st.subheader(f"Forecast Analysis for {selected_bag} at {selected_plant}")
                bag_data = df[(df['Cement Plant Sname'] == selected_plant) & (df['MAKTX'] == selected_bag)]
                month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
                usage_data = []
                for month in month_columns:
                    try:
                        date = pd.to_datetime(month)
                        usage = bag_data[month].iloc[0]
                        usage_data.append({'Date': date, 'Usage': usage})
                    except:
                        continue
                usage_df = pd.DataFrame(usage_data)
                usage_df = usage_df.sort_values('Date')
                mar_2025_data = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2025-03']['Usage'].iloc[0]
                forecaster = BagDemandForecaster(usage_df)
                forecast_results = forecaster.generate_forecast(mar_2025_data)
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("April 2025 Forecast", f"{forecast_results['point_forecast']:,.0f}")
                with col2:
                    st.metric("March 2025 (Projected)", f"{forecast_results['february_extrapolated']:,.0f}")
                with col3:
                    prev_april = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2024-04']['Usage'].iloc[0]
                    yoy_change = ((forecast_results['point_forecast'] - prev_april) / prev_april * 100)
                    st.metric("Year-over-Year Change",f"{yoy_change:,.1f}%")
                fig = go.Figure()
                fig.add_trace(go.Scatter(x=usage_df['Date'],y=usage_df['Usage'],name='Historical',line=dict(color='#2E86C1', width=2)))
                mar_date = pd.to_datetime('2025-03-01')
                fig.add_trace(go.Scatter(x=[feb_date],y=[forecast_results['march_extrapolated']],name='Mar Projection',mode='markers',marker=dict(color='orange', size=10, symbol='diamond')))
                april_date = pd.to_datetime('2025-04-01')
                fig.add_trace(go.Scatter(x=[april_date],y=[forecast_results['point_forecast']],name='April Forecast',mode='markers',marker=dict(color='red', size=12, symbol='star')))
                fig.add_trace(go.Scatter(x=[april_date, april_date],y=[forecast_results['lower_bound'],forecast_results['upper_bound']],name='95% Confidence',mode='lines',line=dict(color='rgba(255,0,0,0.2)', width=10)))
                fig.update_layout(title='Historical Data and Forecast',xaxis_title='Date',yaxis_title='Usage',showlegend=True,hovermode='x unified')
                st.plotly_chart(fig, use_container_width=True)
                st.subheader("Forecasting Model Details")
                individual_forecasts = forecast_results['individual_forecasts']
                for model, value in individual_forecasts.items():
                    st.write(f"**{model.replace('_', ' ').title()}**: {value:,.0f}")
                st.write(f"**Confidence Interval**: {forecast_results['lower_bound']:,.0f} - {forecast_results['upper_bound']:,.0f}")
        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.write("Please ensure your data file has the correct format and try again.")
if __name__ == "__main__":
    main()
