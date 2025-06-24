import pandas as pd
import numpy as np
import warnings
warnings.filterwarnings('ignore')
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import StandardScaler, MinMaxScaler
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import xgboost as xgb
import tensorflow as tf
from tensorflow.keras.models import Sequential
from tensorflow.keras.layers import LSTM, Dense, Dropout
from tensorflow.keras.optimizers import Adam
from tensorflow.keras.callbacks import EarlyStopping
from prophet import Prophet
from datetime import datetime, timedelta
import pickle
import os
class DistrictChannelMaterialForecastingSystem:
    def __init__(self, clustering_results_file, original_data_file):
        self.clustering_results_file = clustering_results_file
        self.original_data_file = original_data_file
        self.cluster_assignments = None
        self.sales_data = None
        self.combination_data = {}
        self.trained_models = {}
        self.forecasts = {}
        self.scalers = {}
    def clean_data(self, data):
        data = data.replace([np.inf, -np.inf], np.nan)
        for col in data.select_dtypes(include=[np.number]).columns:
            Q1 = data[col].quantile(0.25)
            Q3 = data[col].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 3 * IQR
            upper_bound = Q3 + 3 * IQR
            data[col] = data[col].clip(lower=max(lower_bound, -1e6), upper=min(upper_bound, 1e6))
        data = data.fillna(method='bfill').fillna(method='ffill').fillna(0)
        return data
    def validate_numeric_data(self, data, data_name=""):
        if isinstance(data, pd.DataFrame):
            numeric_cols = data.select_dtypes(include=[np.number]).columns
            for col in numeric_cols:
                if data[col].isnull().any():
                    print(f"Warning: {data_name} contains NaN values in column {col}")
                if np.isinf(data[col]).any():
                    print(f"Warning: {data_name} contains infinity values in column {col}")
                if (np.abs(data[col]) > 1e6).any():
                    print(f"Warning: {data_name} contains very large values in column {col}")
        elif isinstance(data, np.ndarray):
            if np.isnan(data).any():
                print(f"Warning: {data_name} contains NaN values")
            if np.isinf(data).any():
                print(f"Warning: {data_name} contains infinity values")
            if (np.abs(data) > 1e6).any():
                print(f"Warning: {data_name} contains very large values")
        return True
    def load_data(self):
        print("Loading data...")
        try:
            self.cluster_assignments = pd.read_excel(self.clustering_results_file, sheet_name='District_Assignments')
            self.sales_data = pd.read_excel(self.original_data_file)
            self.cluster_assignments = self.clean_data(self.cluster_assignments)
            self.sales_data = self.clean_data(self.sales_data)
            print(f"Cluster assignments loaded: {len(self.cluster_assignments)} districts")
            print(f"Sales data loaded: {self.sales_data.shape}")
            return True
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            return False
    def prepare_combination_data(self):
        print("Preparing district-channel-material combination data...")
        time_columns = ['Jan\'23', 'Feb\'23', 'Mar\'23', 'Apr\'23', 'May\'23', 'Jun\'23','Jul\'23', 'Aug\'23', 'Sep\'23', 'Oct\'23', 'Nov\'23', 'Dec\'23','Jan\'24', 'Feb\'24', 'Mar\'24', 'Apr\'24', 'May\'24', 'Jun\'24','Jul\'24', 'Aug\'24', 'Sep\'24', 'Oct\'24', 'Nov\'24', 'Dec\'24']
        dates = []
        for col in time_columns:
            if '23' in col:
                year = '2023'
            else:
                year = '2024'
            month = col.split('\'')[0]
            month_num = {'Jan': '01', 'Feb': '02', 'Mar': '03', 'Apr': '04','May': '05', 'Jun': '06', 'Jul': '07', 'Aug': '08','Sep': '09', 'Oct': '10', 'Nov': '11', 'Dec': '12'}[month]
            dates.append(f"{year}-{month_num}-01")
        district_cluster_map = {}
        for _, row in self.cluster_assignments.iterrows():
            district = row['District']
            cluster = row['Cluster']
            if cluster != 'Noise':
                district_cluster_map[district] = cluster
        for _, row in self.sales_data.iterrows():
            district = row['District Name']
            channel = row['DISTRIBUTION CHANNEL']
            material = row['MATERIAL TYPE']
            if district not in district_cluster_map:
                continue
            cluster = district_cluster_map[district]
            combination_key = f"{district}_{channel}_{material}_{cluster}"
            print(f"Processing: {district} | {channel} | {material} (Cluster: {cluster})")
            monthly_sales = []
            for col in time_columns:
                if col in row.index:
                    sales_value = pd.to_numeric(row[col], errors='coerce')
                    if pd.isna(sales_value) or np.isinf(sales_value) or abs(sales_value) > 1e6:
                        sales_value = 0
                    monthly_sales.append(max(0, sales_value))
                else:
                    monthly_sales.append(0)
            ts_data = pd.DataFrame({'ds': pd.to_datetime(dates),'y': monthly_sales})
            ts_data['month'] = ts_data['ds'].dt.month
            ts_data['quarter'] = ts_data['ds'].dt.quarter
            ts_data['year'] = ts_data['ds'].dt.year
            ts_data['month_sin'] = np.sin(2 * np.pi * ts_data['month'] / 12)
            ts_data['month_cos'] = np.cos(2 * np.pi * ts_data['month'] / 12)
            for lag in [1, 2, 3, 6, 12]:
                if lag < len(ts_data):
                    ts_data[f'lag_{lag}'] = ts_data['y'].shift(lag)
            ts_data['rolling_mean_3'] = ts_data['y'].rolling(window=3, min_periods=1).mean()
            ts_data['rolling_std_3'] = ts_data['y'].rolling(window=3, min_periods=1).std()
            ts_data['rolling_mean_6'] = ts_data['y'].rolling(window=6, min_periods=1).mean()
            ts_data['trend'] = range(len(ts_data))
            yoy_change = ts_data['y'].pct_change(periods=12)
            yoy_change = yoy_change.clip(-0.9, 10.0)
            ts_data['yoy_growth'] = yoy_change
            ts_data = self.clean_data(ts_data)
            self.validate_numeric_data(ts_data, f"Combination {combination_key}")
            self.combination_data[combination_key] = {'data': ts_data,'district': district,'channel': channel,'material': material,'cluster': cluster}
            print(f"Combination {combination_key}: Sales range: {ts_data['y'].min():.0f} - {ts_data['y'].max():.0f}")
    def create_lstm_sequences(self, data, sequence_length=6):
        X, y = [], []
        for i in range(sequence_length, len(data)):
            X.append(data[i-sequence_length:i])
            y.append(data[i])
        return np.array(X), np.array(y)
    def train_lstm_model(self, combination_key, sequence_length=6):
        district = self.combination_data[combination_key]['district']
        channel = self.combination_data[combination_key]['channel']
        material = self.combination_data[combination_key]['material']
        print(f"Training LSTM for {district} | {channel} | {material}...")
        ts_data = self.combination_data[combination_key]['data'].copy()
        if len(ts_data) < sequence_length + 2:
            print(f"Insufficient data for LSTM training in combination {combination_key}")
            return None
        try:
            sales_values = ts_data[['y']].values
            sales_values = np.nan_to_num(sales_values, nan=0.0, posinf=1e6, neginf=0.0)
            scaler = MinMaxScaler(feature_range=(0.1, 0.9))
            scaled_data = scaler.fit_transform(sales_values)
            if np.isnan(scaled_data).any() or np.isinf(scaled_data).any():
                print(f"Data scaling failed for combination {combination_key}")
                return None
            X, y = self.create_lstm_sequences(scaled_data.flatten(), sequence_length)
            if len(X) < 3:
                print(f"Insufficient sequences for LSTM training in combination {combination_key}")
                return None
            X = X.reshape((X.shape[0], X.shape[1], 1))
            if np.isnan(X).any() or np.isinf(X).any() or np.isnan(y).any() or np.isinf(y).any():
                print(f"Invalid training data for combination {combination_key}")
                return None
            model = Sequential([LSTM(16, activation='tanh', input_shape=(sequence_length, 1), return_sequences=False),Dropout(0.1),Dense(8, activation='relu'),Dense(1, activation='sigmoid')])
            model.compile(optimizer=Adam(learning_rate=0.001), loss='mse', metrics=['mae'])
            early_stop = EarlyStopping(monitor='loss', patience=10, restore_best_weights=True)
            history = model.fit(X, y,epochs=30,batch_size=min(4, len(X)),callbacks=[early_stop],verbose=0,validation_split=0.2 if len(X) > 5 else 0)
            self.scalers[f'{combination_key}_lstm'] = scaler
            return model
        except Exception as e:
            print(f"LSTM training failed for combination {combination_key}: {str(e)}")
            return None
    def train_random_forest_model(self, combination_key):
        district = self.combination_data[combination_key]['district']
        channel = self.combination_data[combination_key]['channel']
        material = self.combination_data[combination_key]['material']
        print(f"Training Random Forest for {district} | {channel} | {material}...")
        try:
            ts_data = self.combination_data[combination_key]['data'].copy()
            feature_cols = [col for col in ts_data.columns if col not in ['ds', 'y']]
            X = ts_data[feature_cols]
            y = ts_data['y']
            X = self.clean_data(X)
            y = np.nan_to_num(y, nan=0.0, posinf=1e6, neginf=0.0)
            self.validate_numeric_data(X, f"Combination {combination_key} RF features")
            if len(X) < 3:
                print(f"Insufficient data for Random Forest training in combination {combination_key}")
                return None
            model = RandomForestRegressor(n_estimators=50,max_depth=5,min_samples_split=2,min_samples_leaf=1,random_state=42,n_jobs=1)
            model.fit(X, y)
            return model
        except Exception as e:
            print(f"Random Forest training failed for combination {combination_key}: {str(e)}")
            return None
    def train_xgboost_model(self, combination_key):
        district = self.combination_data[combination_key]['district']
        channel = self.combination_data[combination_key]['channel']
        material = self.combination_data[combination_key]['material']
        print(f"Training XGBoost for {district} | {channel} | {material}...")
        try:
            ts_data = self.combination_data[combination_key]['data'].copy()
            feature_cols = [col for col in ts_data.columns if col not in ['ds', 'y']]
            X = ts_data[feature_cols]
            y = ts_data['y']
            X = self.clean_data(X)
            y = np.nan_to_num(y, nan=0.0, posinf=1e6, neginf=0.0)
            if len(X) < 3:
                print(f"Insufficient data for XGBoost training in combination {combination_key}")
                return None
            model = xgb.XGBRegressor(n_estimators=50,max_depth=3,learning_rate=0.1,subsample=0.8,colsample_bytree=0.8,random_state=42,reg_alpha=0.1,reg_lambda=0.1,n_jobs=1)
            model.fit(X, y)
            return model
        except Exception as e:
            print(f"XGBoost training failed for combination {combination_key}: {str(e)}")
            return None
    def train_prophet_model(self, combination_key):
        district = self.combination_data[combination_key]['district']
        channel = self.combination_data[combination_key]['channel']
        material = self.combination_data[combination_key]['material']
        print(f"Training Prophet for {district} | {channel} | {material}...")
        try:
            ts_data = self.combination_data[combination_key]['data'].copy()
            if len(ts_data) < 10:
                print(f"Insufficient data for Prophet training in combination {combination_key}")
                return None
            prophet_data = ts_data[['ds', 'y']].copy()
            prophet_data['y'] = np.nan_to_num(prophet_data['y'], nan=0.0, posinf=1e6, neginf=0.0)
            prophet_data['y'] = np.maximum(prophet_data['y'], 0)
            model = Prophet(yearly_seasonality=True,weekly_seasonality=False,daily_seasonality=False,changepoint_prior_scale=0.05,seasonality_prior_scale=10,mcmc_samples=0,uncertainty_samples=100)
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                model.fit(prophet_data)
            return model
        except Exception as e:
            print(f"Prophet training failed for combination {combination_key}: {str(e)}")
            return None
    def train_all_models(self):
        print("Training all models for all district-channel-material combinations...")
        print("=" * 80)
        self.trained_models = {}
        for combination_key in self.combination_data.keys():
            district = self.combination_data[combination_key]['district']
            channel = self.combination_data[combination_key]['channel']
            material = self.combination_data[combination_key]['material']
            cluster = self.combination_data[combination_key]['cluster']
            print(f"\nTraining models for {district} | {channel} | {material} (Cluster: {cluster})...")
            combination_models = {}
            lstm_model = self.train_lstm_model(combination_key)
            if lstm_model:
                combination_models['LSTM'] = lstm_model
            rf_model = self.train_random_forest_model(combination_key)
            if rf_model:
                combination_models['RandomForest'] = rf_model
            xgb_model = self.train_xgboost_model(combination_key)
            if xgb_model:
                combination_models['XGBoost'] = xgb_model
            prophet_model = self.train_prophet_model(combination_key)
            if prophet_model:
                combination_models['Prophet'] = prophet_model
            self.trained_models[combination_key] = combination_models
            print(f"Completed training for {district} | {channel} | {material}: {len(combination_models)} models trained")
    def generate_forecasts(self, forecast_months=['2025-01-01', '2025-02-01', '2025-03-01']):
        print("Generating district-channel-material level forecasts for Jan'25 to Mar'25...")
        self.forecasts = {}
        forecast_dates = pd.to_datetime(forecast_months)
        for combination_key in self.combination_data.keys():
            district = self.combination_data[combination_key]['district']
            channel = self.combination_data[combination_key]['channel']
            material = self.combination_data[combination_key]['material']
            cluster = self.combination_data[combination_key]['cluster']
            print(f"Forecasting for {district} | {channel} | {material} (Cluster: {cluster})...")
            combination_forecasts = {}
            ts_data = self.combination_data[combination_key]['data'].copy()
            baseline_forecast = max(ts_data['y'].mean(), 0)
            if 'LSTM' in self.trained_models[combination_key]:
                try:
                    lstm_forecasts = []
                    model = self.trained_models[combination_key]['LSTM']
                    scaler = self.scalers[f'{combination_key}_lstm']
                    last_sequence = scaler.transform(ts_data[['y']].tail(6).values).flatten()
                    for _ in range(3):
                        X_pred = last_sequence[-6:].reshape(1, 6, 1)
                        pred_scaled = model.predict(X_pred, verbose=0)[0][0]
                        if np.isnan(pred_scaled) or np.isinf(pred_scaled):
                            pred_scaled = 0.5
                        pred_actual = scaler.inverse_transform([[pred_scaled]])[0][0]
                        if np.isnan(pred_actual) or np.isinf(pred_actual) or pred_actual < 0:
                            pred_actual = baseline_forecast
                        lstm_forecasts.append(pred_actual)
                        last_sequence = np.append(last_sequence, pred_scaled)
                    combination_forecasts['LSTM'] = lstm_forecasts
                except Exception as e:
                    print(f"LSTM forecasting failed for combination {combination_key}: {str(e)}")
                    combination_forecasts['LSTM'] = [baseline_forecast] * 3
            if 'RandomForest' in self.trained_models[combination_key]:
                try:
                    rf_forecasts = []
                    model = self.trained_models[combination_key]['RandomForest']
                    for i, date in enumerate(forecast_dates):
                        features = {'month': date.month,'quarter': date.quarter,'year': date.year,'month_sin': np.sin(2 * np.pi * date.month / 12),'month_cos': np.cos(2 * np.pi * date.month / 12),'trend': len(ts_data) + i}
                        last_values = ts_data['y'].tolist()
                        if i > 0:
                            last_values.extend(rf_forecasts)
                        for lag in [1, 2, 3, 6, 12]:
                            if len(last_values) >= lag:
                                features[f'lag_{lag}'] = last_values[-lag]
                            else:
                                features[f'lag_{lag}'] = baseline_forecast
                        recent_values = last_values[-6:] if len(last_values) >= 6 else last_values
                        features['rolling_mean_3'] = np.mean(recent_values[-3:]) if len(recent_values) >= 3 else baseline_forecast
                        features['rolling_std_3'] = np.std(recent_values[-3:]) if len(recent_values) >= 3 else 0
                        features['rolling_mean_6'] = np.mean(recent_values) if len(recent_values) > 0 else baseline_forecast
                        features['yoy_growth'] = 0
                        feature_cols = [col for col in ts_data.columns if col not in ['ds', 'y']]
                        X_pred = []
                        for col in feature_cols:
                            val = features.get(col, 0)
                            if np.isnan(val) or np.isinf(val):
                                val = 0
                            X_pred.append(val)
                        pred = model.predict([X_pred])[0]
                        if np.isnan(pred) or np.isinf(pred) or pred < 0:
                            pred = baseline_forecast
                        rf_forecasts.append(pred)
                    combination_forecasts['RandomForest'] = rf_forecasts
                except Exception as e:
                    print(f"Random Forest forecasting failed for combination {combination_key}: {str(e)}")
                    combination_forecasts['RandomForest'] = [baseline_forecast] * 3
            if 'XGBoost' in self.trained_models[combination_key]:
                try:
                    xgb_forecasts = []
                    model = self.trained_models[combination_key]['XGBoost']
                    for i, date in enumerate(forecast_dates):
                        features = {'month': date.month,'quarter': date.quarter,'year': date.year,'month_sin': np.sin(2 * np.pi * date.month / 12),'month_cos': np.cos(2 * np.pi * date.month / 12),'trend': len(ts_data) + i}
                        last_values = ts_data['y'].tolist()
                        if i > 0:
                            last_values.extend(xgb_forecasts)
                        for lag in [1, 2, 3, 6, 12]:
                            if len(last_values) >= lag:
                                features[f'lag_{lag}'] = last_values[-lag]
                            else:
                                features[f'lag_{lag}'] = baseline_forecast
                        recent_values = last_values[-6:] if len(last_values) >= 6 else last_values
                        features['rolling_mean_3'] = np.mean(recent_values[-3:]) if len(recent_values) >= 3 else baseline_forecast
                        features['rolling_std_3'] = np.std(recent_values[-3:]) if len(recent_values) >= 3 else 0
                        features['rolling_mean_6'] = np.mean(recent_values) if len(recent_values) > 0 else baseline_forecast
                        features['yoy_growth'] = 0
                        feature_cols = [col for col in ts_data.columns if col not in ['ds', 'y']]
                        X_pred = []
                        for col in feature_cols:
                            val = features.get(col, 0)
                            if np.isnan(val) or np.isinf(val):
                                val = 0
                            X_pred.append(val)
                        pred = model.predict([X_pred])[0]
                        if np.isnan(pred) or np.isinf(pred) or pred < 0:
                            pred = baseline_forecast
                        xgb_forecasts.append(pred)
                    combination_forecasts['XGBoost'] = xgb_forecasts
                except Exception as e:
                    print(f"XGBoost forecasting failed for combination {combination_key}: {str(e)}")
                    combination_forecasts['XGBoost'] = [baseline_forecast] * 3
            if 'Prophet' in self.trained_models[combination_key]:
                try:
                    model = self.trained_models[combination_key]['Prophet']
                    future = pd.DataFrame({'ds': forecast_dates})
                    with warnings.catch_warnings():
                        warnings.simplefilter("ignore")
                        forecast = model.predict(future)
                    prophet_forecasts = []
                    for val in forecast['yhat'].tolist():
                        if np.isnan(val) or np.isinf(val) or val < 0:
                            val = baseline_forecast
                        prophet_forecasts.append(val)
                    combination_forecasts['Prophet'] = prophet_forecasts
                except Exception as e:
                    print(f"Prophet forecasting failed for combination {combination_key}: {str(e)}")
                    combination_forecasts['Prophet'] = [baseline_forecast] * 3
            if not combination_forecasts:
                combination_forecasts['Baseline'] = [baseline_forecast] * 3
            self.forecasts[combination_key] = combination_forecasts
    def create_ensemble_forecast(self):
        print("Creating ensemble forecasts for each district-channel-material combination...")
        for combination_key in self.forecasts.keys():
            combination_forecasts = self.forecasts[combination_key]
            if len(combination_forecasts) == 0:
                continue
            ensemble_forecasts = []
            for month_idx in range(3):
                month_predictions = []
                for model_name, predictions in combination_forecasts.items():
                    if len(predictions) > month_idx:
                        pred = predictions[month_idx]
                        if not (np.isnan(pred) or np.isinf(pred)) and pred >= 0:
                            month_predictions.append(pred)
                if month_predictions:
                    ensemble_forecast = np.mean(month_predictions)
                    if np.isnan(ensemble_forecast) or np.isinf(ensemble_forecast) or ensemble_forecast < 0:
                        ensemble_forecast = 0
                    ensemble_forecasts.append(ensemble_forecast)
                else:
                    ensemble_forecasts.append(0)
            self.forecasts[combination_key]['Ensemble'] = ensemble_forecasts
    def export_forecasts(self, output_file='district_channel_material_forecasts.xlsx'):
        print(f"Exporting district-channel-material level forecasts to {output_file}...")
        forecast_results = []
        combination_summary = []
        district_summary = {}
        channel_summary = {}
        material_summary = {}
        cluster_summary = {}
        months = ['Jan_2025', 'Feb_2025', 'Mar_2025']
        for combination_key in self.forecasts.keys():
            district = self.combination_data[combination_key]['district']
            channel = self.combination_data[combination_key]['channel']
            material = self.combination_data[combination_key]['material']
            cluster = self.combination_data[combination_key]['cluster']
            combination_forecasts = self.forecasts[combination_key]
            for model_name, predictions in combination_forecasts.items():
                row = {'District': district,'Channel': channel,'Material_Type': material,'Cluster': cluster,'Model': model_name,}
                for i, month in enumerate(months):
                    if i < len(predictions):
                        row[month] = predictions[i]
                    else:
                        row[month] = 0
                row['Total_Forecast'] = sum(predictions)
                forecast_results.append(row)
            if 'Ensemble' in combination_forecasts:
                ensemble_forecast = combination_forecasts['Ensemble']
                combination_summary.append({'District': district,'Channel': channel,'Material_Type': material,'Cluster': cluster,'Jan_2025_Forecast': ensemble_forecast[0] if len(ensemble_forecast) > 0 else 0,'Feb_2025_Forecast': ensemble_forecast[1] if len(ensemble_forecast) > 1 else 0,'Mar_2025_Forecast': ensemble_forecast[2] if len(ensemble_forecast) > 2 else 0,'Total_Q1_2025_Forecast': sum(ensemble_forecast)})
                if district not in district_summary:
                    district_summary[district] = {'Jan': 0, 'Feb': 0, 'Mar': 0, 'Total': 0}
                district_summary[district]['Jan'] += ensemble_forecast[0] if len(ensemble_forecast) > 0 else 0
                district_summary[district]['Feb'] += ensemble_forecast[1] if len(ensemble_forecast) > 1 else 0
                district_summary[district]['Mar'] += ensemble_forecast[2] if len(ensemble_forecast) > 2 else 0
                district_summary[district]['Total'] += sum(ensemble_forecast)
                if channel not in channel_summary:
                    channel_summary[channel] = {'Jan': 0, 'Feb': 0, 'Mar': 0, 'Total': 0}
                channel_summary[channel]['Jan'] += ensemble_forecast[0] if len(ensemble_forecast) > 0 else 0
                channel_summary[channel]['Feb'] += ensemble_forecast[1] if len(ensemble_forecast) > 1 else 0
                channel_summary[channel]['Mar'] += ensemble_forecast[2] if len(ensemble_forecast) > 2 else 0
                channel_summary[channel]['Total'] += sum(ensemble_forecast)
                if material not in material_summary:
                    material_summary[material] = {'Jan': 0, 'Feb': 0, 'Mar': 0, 'Total': 0}
                material_summary[material]['Jan'] += ensemble_forecast[0] if len(ensemble_forecast) > 0 else 0
                material_summary[material]['Feb'] += ensemble_forecast[1] if len(ensemble_forecast) > 1 else 0
                material_summary[material]['Mar'] += ensemble_forecast[2] if len(ensemble_forecast) > 2 else 0
                material_summary[material]['Total'] += sum(ensemble_forecast)
                if cluster not in cluster_summary:
                    cluster_summary[cluster] = {'Jan': 0, 'Feb': 0, 'Mar': 0, 'Total': 0}
                cluster_summary[cluster]['Jan'] += ensemble_forecast[0] if len(ensemble_forecast) > 0 else 0
                cluster_summary[cluster]['Feb'] += ensemble_forecast[1] if len(ensemble_forecast) > 1 else 0
                cluster_summary[cluster]['Mar'] += ensemble_forecast[2] if len(ensemble_forecast) > 2 else 0
                cluster_summary[cluster]['Total'] += sum(ensemble_forecast)
        district_summary_df = pd.DataFrame([{'District': district,'Jan_2025_Forecast': data['Jan'],'Feb_2025_Forecast': data['Feb'],'Mar_2025_Forecast': data['Mar'],'Total_Q1_2025_Forecast': data['Total']}for district, data in district_summary.items()])
        channel_summary_df = pd.DataFrame([{'Channel': channel,'Jan_2025_Forecast': data['Jan'],'Feb_2025_Forecast': data['Feb'],'Mar_2025_Forecast': data['Mar'],'Total_Q1_2025_Forecast': data['Total']}for channel, data in channel_summary.items()])
        material_summary_df = pd.DataFrame([{'Material_Type': material,'Jan_2025_Forecast': data['Jan'],'Feb_2025_Forecast': data['Feb'],'Mar_2025_Forecast': data['Mar'],'Total_Q1_2025_Forecast': data['Total']}for material, data in material_summary.items()])
        cluster_summary_df = pd.DataFrame([{'Cluster': cluster,'Jan_2025_Forecast': data['Jan'],'Feb_2025_Forecast': data['Feb'],'Mar_2025_Forecast': data['Mar'],'Total_Q1_2025_Forecast': data['Total']}for cluster, data in cluster_summary.items()])
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            pd.DataFrame(forecast_results).to_excel(writer, sheet_name='All_Model_Forecasts', index=False)
            pd.DataFrame(combination_summary).to_excel(writer, sheet_name='Combination_Forecasts', index=False)
            district_summary_df.to_excel(writer, sheet_name='District_Summary', index=False)
            channel_summary_df.to_excel(writer, sheet_name='Channel_Summary', index=False)
            material_summary_df.to_excel(writer, sheet_name='Material_Summary', index=False)
            cluster_summary_df.to_excel(writer, sheet_name='Cluster_Summary', index=False)
        print(f"Forecasts exported successfully to {output_file}")
        print(f"Total combinations forecasted: {len(combination_summary)}")
        print(f"Total districts: {len(district_summary)}")
        print(f"Total channels: {len(channel_summary)}")
        print(f"Total materials: {len(material_summary)}")
        print(f"Total clusters: {len(cluster_summary)}")
    def save_models(self, models_dir='trained_models'):
        if not os.path.exists(models_dir):
            os.makedirs(models_dir)
        print(f"Saving trained models to {models_dir}...")
        for combination_key, models in self.trained_models.items():
            combination_dir = os.path.join(models_dir, combination_key.replace('/', '_').replace('\\', '_'))
            if not os.path.exists(combination_dir):
                os.makedirs(combination_dir)
            for model_name, model in models.items():
                model_path = os.path.join(combination_dir, f"{model_name}.pkl")
                try:
                    if model_name == 'LSTM':
                        model.save(os.path.join(combination_dir, f"{model_name}.h5"))
                    else:
                        with open(model_path, 'wb') as f:
                            pickle.dump(model, f)
                    print(f"Saved {model_name} model for {combination_key}")
                except Exception as e:
                    print(f"Failed to save {model_name} model for {combination_key}: {str(e)}")
        scalers_path = os.path.join(models_dir, 'scalers.pkl')
        with open(scalers_path, 'wb') as f:
            pickle.dump(self.scalers, f)
        print("Model saving completed!")
    def run_complete_forecasting_pipeline(self, output_file='district_channel_material_forecasts.xlsx'):
        print("Starting District-Channel-Material Level Forecasting Pipeline...")
        print("=" * 80)
        if not self.load_data():
            print("Failed to load data. Exiting...")
            return False
        self.prepare_combination_data()
        print(f"Prepared data for {len(self.combination_data)} district-channel-material combinations")
        self.train_all_models()
        self.generate_forecasts()
        self.create_ensemble_forecast()
        self.export_forecasts(output_file)
        self.save_models()
        print("=" * 80)
        print("Forecasting pipeline completed successfully!")
        print(f"Results saved to: {output_file}")
        return True
    def get_forecast_summary(self):
        if not self.forecasts:
            print("No forecasts available. Run the forecasting pipeline first.")
            return None
        summary = {'total_combinations': len(self.forecasts),'districts': set(),'channels': set(),'materials': set(),'clusters': set(),'total_forecast_q1_2025': 0}
        for combination_key in self.forecasts.keys():
            combination_info = self.combination_data[combination_key]
            summary['districts'].add(combination_info['district'])
            summary['channels'].add(combination_info['channel'])
            summary['materials'].add(combination_info['material'])
            summary['clusters'].add(combination_info['cluster'])
            if 'Ensemble' in self.forecasts[combination_key]:
                ensemble_forecast = self.forecasts[combination_key]['Ensemble']
                summary['total_forecast_q1_2025'] += sum(ensemble_forecast)
        summary['districts'] = len(summary['districts'])
        summary['channels'] = len(summary['channels'])
        summary['materials'] = len(summary['materials'])
        summary['clusters'] = len(summary['clusters'])
        return summary
if __name__ == "__main__":
    forecasting_system = DistrictChannelMaterialForecastingSystem(clustering_results_file='/content/optimal_district_clustering.xlsx',original_data_file='/content/Forecasting Data.xlsx')
    success = forecasting_system.run_complete_forecasting_pipeline(output_file='district_channel_material_forecasts_2025.xlsx')
    if success:
        summary = forecasting_system.get_forecast_summary()
        if summary:
            print("\nForecast Summary:")
            print(f"- Total combinations: {summary['total_combinations']}")
            print(f"- Unique districts: {summary['districts']}")
            print(f"- Unique channels: {summary['channels']}")
            print(f"- Unique materials: {summary['materials']}")
            print(f"- Unique clusters: {summary['clusters']}")
            print(f"- Total Q1 2025 forecast: {summary['total_forecast_q1_2025']:,.0f}")
    else:
        print("Forecasting pipeline failed!")
