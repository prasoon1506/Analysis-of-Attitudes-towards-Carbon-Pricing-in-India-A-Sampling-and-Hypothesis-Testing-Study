import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.seasonal import seasonal_decompose
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import TimeSeriesSplit
from sklearn.metrics import mean_squared_error, mean_absolute_percentage_error
import plotly.graph_objects as go
from datetime import datetime
import warnings
from scipy.optimize import minimize
import io
from google.colab import files
warnings.filterwarnings('ignore')

class ClusterEnsembleOptimizer:
    """Optimizes ensemble weights for forecasting based on cluster membership."""
    
    def __init__(self, cluster_data, historical_data):
        """
        Initialize the optimizer with cluster assignments and historical data.
        
        Args:
            cluster_data: DataFrame with bag cluster assignments
            historical_data: DataFrame with usage time series data
        """
        self.cluster_data = cluster_data
        self.historical_data = historical_data
        self.clusters = {}
        self.optimal_weights = {}
        self.results = {}
        self.forecasts = {}
        
    def _prepare_time_series(self, plant, bag):
        """Extract and prepare time series data for a specific plant-bag combination."""
        bag_data = self.historical_data[(self.historical_data['Cement Plant Sname'] == plant) & 
                                       (self.historical_data['MAKTX'] == bag)]
        
        month_columns = [col for col in self.historical_data.columns 
                       if col not in ['Cement Plant Sname', 'MAKTX']]
        
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
        usage_df['Year'] = usage_df['Date'].dt.year
        usage_df['Month'] = usage_df['Date'].dt.month
        
        return usage_df
    
    def organize_by_cluster(self):
        """Group bags by their assigned clusters."""
        # Get unique clusters
        unique_clusters = self.cluster_data['Cluster'].unique()
        
        for cluster in unique_clusters:
            # Filter items in this cluster
            cluster_items = self.cluster_data[self.cluster_data['Cluster'] == cluster]
            
            # Initialize cluster data structure
            self.clusters[cluster] = {'items': [], 'data': []}
            
            # Process each item in the cluster
            for _, item in cluster_items.iterrows():
                plant = item['Plant']
                bag = item['Bag']
                
                try:
                    # Get time series for this plant-bag
                    usage_df = self._prepare_time_series(plant, bag)
                    
                    if len(usage_df) >= 12:  # Only include if we have at least 12 months of data
                        self.clusters[cluster]['items'].append({
                            'plant': plant,
                            'bag': bag
                        })
                        self.clusters[cluster]['data'].append(usage_df)
                except Exception as e:
                    print(f"Error processing {plant}-{bag}: {str(e)}")
        
        # Print summary
        for cluster, data in self.clusters.items():
            print(f"{cluster}: {len(data['items'])} items with sufficient data")
    
    def _calculate_seasonal_indices(self, time_series):
        """Calculate seasonal indices for the given time series."""
        try:
            decomposition = seasonal_decompose(time_series['Usage'], period=12, model='multiplicative')
            seasonal = pd.Series(decomposition.seasonal)
            monthly_indices = {}
            
            for month in range(1, 13):
                month_indices = time_series['Month'] == month
                if any(month_indices):
                    monthly_indices[month] = seasonal[month_indices].mean()
                else:
                    monthly_indices[month] = 1.0
                    
            return monthly_indices
        except:
            return {m: 1.0 for m in range(1, 13)}
    
    def _extrapolate_february(self, feb_partial_data):
        """Extrapolate full February data from partial data."""
        days_in_feb = 28  # Adjust for leap years as needed
        current_days = 16  # Assuming data is from first 16 days
        daily_rate = feb_partial_data / current_days
        return daily_rate * days_in_feb
    
    def _get_recent_trend(self, time_series, window=6):
        """Calculate recent trend from the last few months."""
        recent_data = time_series.tail(window)
        if len(recent_data) > 1:
            return (recent_data['Usage'].iloc[-1] - recent_data['Usage'].iloc[0]) / (window - 1)
        return 0
    
    def _generate_individual_forecasts(self, time_series, feb_data):
        """Generate forecasts using individual models."""
        # Use February data directly
        feb_full = feb_data
        
        # Calculate seasonal factors
        seasonal_indices = self._calculate_seasonal_indices(time_series)
        march_seasonal_factor = seasonal_indices.get(3, 1.0)
        
        # Holt-Winters Forecasting
        try:
            model_hw = ExponentialSmoothing(
                time_series['Usage'],
                seasonal_periods=12,
                trend='add',
                seasonal='mul'
            ).fit()
            hw_forecast = model_hw.forecast(1)[0]
        except:
            hw_forecast = time_series['Usage'].mean() * march_seasonal_factor
        
        # Random Forest
        try:
            rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
            X = pd.DataFrame({
                'Month': time_series['Month'],
                'Year': time_series['Year'],
                'Lag1': time_series['Usage'].shift(1),
                'Lag12': time_series['Usage'].shift(12),
                'Trend': range(len(time_series))
            }).dropna()
            
            y = time_series['Usage'].iloc[12:]
            rf_model.fit(X, y)
            
            march_features = pd.DataFrame({
                'Month': [3],
                'Year': [2025],
                'Lag1': [feb_full],
                'Lag12': [time_series['Usage'].iloc[-12] if len(time_series) >= 12 else time_series['Usage'].mean()],
                'Trend': [len(time_series)]
            })
            
            rf_forecast = rf_model.predict(march_features)[0]
        except:
            rf_forecast = time_series['Usage'].mean() * march_seasonal_factor
        
        # Trend-based forecast
        trend = self._get_recent_trend(time_series)
        trend_forecast = feb_full + trend
        
        return {
            'holt_winters': hw_forecast,
            'random_forest': rf_forecast,
            'trend_based': trend_forecast,
            'february_value': feb_full
        }
    def _objective_function(self, weights, forecasts, actuals):
     weights = np.array(weights) / sum(weights)  # Normalize weights to sum to 1
    
     combined_forecasts = np.zeros(len(actuals))
     for i in range(len(actuals)):
        combined_forecasts[i] = (
            weights[0] * forecasts[i]['holt_winters'] +
            weights[1] * forecasts[i]['random_forest'] +
            weights[2] * forecasts[i]['trend_based']
        )
    
    # Filter out zeros and NaN values in actuals to avoid division by zero
     valid_indices = (actuals != 0) & (~np.isnan(actuals)) & (~np.isnan(combined_forecasts))
    
     if sum(valid_indices) == 0:
        # If no valid data points, return a high error value
        return 1000.0
    
    # Calculate MAPE (Mean Absolute Percentage Error) only on valid data points
     valid_actuals = actuals[valid_indices]
     valid_forecasts = combined_forecasts[valid_indices]
    
     mape = np.mean(np.abs((valid_actuals - valid_forecasts) / valid_actuals)) * 100
     return mape

    def _constraints(self):
        """Define constraints for the optimization (weights sum to 1)."""
        return {'type': 'eq', 'fun': lambda x: sum(x) - 1}
    
    def _bounds(self):
        """Define bounds for the weights (between 0 and 1)."""
        return [(0, 1), (0, 1), (0, 1)]
    
    def optimize_cluster_weights(self):
        """Optimize ensemble weights for each cluster."""
        for cluster, cluster_data in self.clusters.items():
            print(f"\nOptimizing weights for {cluster}...")
            
            all_forecasts = []
            all_actuals = []
            
            # For each bag in the cluster, calculate backtested forecasts
            for i, time_series in enumerate(cluster_data['data']):
                if len(time_series) < 12:  # Need at least 12 months of data for validation
                    continue
                    
                plant = cluster_data['items'][i]['plant']
                bag = cluster_data['items'][i]['bag']
                
                # Use the last 3 months for validation
                train_data = time_series.iloc[:-3].copy()
                test_data = time_series.iloc[-3:].copy()
                
                item_forecasts = []
                item_actuals = []
                
                # For each month in the test set, generate a forecast
                for j in range(len(test_data)):
                    forecast_month = test_data.iloc[j]['Date']
                    forecast_usage = test_data.iloc[j]['Usage']
                    
                    # Use data up to the previous month
                    cutoff_date = forecast_month - pd.DateOffset(months=1)
                    train_subset = train_data[train_data['Date'] <= cutoff_date].copy()
                    
                    # Generate forecasts for this month
                    if j > 0:
                        # Use the previous month's actual value as February data
                        prev_month_data = test_data.iloc[j-1]['Usage']
                    else:
                        # Use the last month from training data
                        prev_month_data = train_subset['Usage'].iloc[-1]
                    
                    try:
                        forecasts = self._generate_individual_forecasts(train_subset, prev_month_data)
                        item_forecasts.append(forecasts)
                        item_actuals.append(forecast_usage)
                    except Exception as e:
                        print(f"Error forecasting {plant}-{bag} for {forecast_month}: {str(e)}")
                
                # Add to overall collections
                all_forecasts.extend(item_forecasts)
                all_actuals.extend(item_actuals)
            
            # Skip optimization if insufficient data
            if len(all_actuals) < 5:
                print(f"Insufficient data for {cluster}. Using default weights.")
                self.optimal_weights[cluster] = {
                    'holt_winters': 0.4,
                    'random_forest': 0.4,
                    'trend_based': 0.2
                }
                continue
            
            # Convert actuals to numpy array
            all_actuals = np.array(all_actuals)
            
            # Initial weights (equal)
            initial_weights = [1/3, 1/3, 1/3]
            
            # Optimize weights
            result = minimize(
                self._objective_function, 
                initial_weights,
                args=(all_forecasts, all_actuals),
                method='SLSQP',
                bounds=self._bounds(),
                constraints=self._constraints()
            )
            
            # Store optimized weights
            optimized_weights = result['x'] / sum(result['x'])  # Normalize to sum to 1
            self.optimal_weights[cluster] = {
                'holt_winters': optimized_weights[0],
                'random_forest': optimized_weights[1],
                'trend_based': optimized_weights[2],
                'mape': result['fun']
            }
            
            print(f"Optimized weights for {cluster}:")
            print(f"  Holt-Winters: {optimized_weights[0]:.3f}")
            print(f"  Random Forest: {optimized_weights[1]:.3f}")
            print(f"  Trend-based: {optimized_weights[2]:.3f}")
            print(f"  Validation MAPE: {result['fun']:.2f}%")
    
    def extract_february_data(self):
        """Extract February 2025 data from historical data."""
        feb_data = {}
        
        # Find February 2025 column
        feb_col = None
        for col in self.historical_data.columns:
            try:
                date = pd.to_datetime(col)
                if date.year == 2025 and date.month == 2:
                    feb_col = col
                    break
            except:
                continue
        
        if feb_col is None:
            print("February 2025 data not found in historical data!")
            return feb_data
        
        # Extract data for each plant-bag combination
        for _, row in self.historical_data.iterrows():
            plant = row['Cement Plant Sname']
            bag = row['MAKTX']
            value = row[feb_col]
            
            # Only include non-null values
            if pd.notna(value):
                feb_data[(plant, bag)] = value
        
        print(f"Extracted February 2025 data for {len(feb_data)} plant-bag combinations")
        return feb_data
    
    def generate_forecasts(self):
        """Generate forecasts for March 2025 using optimized weights and February data from historical file."""
        # Extract February 2025 data
        feb_2025_data = self.extract_february_data()
        
        forecasts = []
        
        # Process each cluster
        for cluster, cluster_data in self.clusters.items():
            # Get optimal weights for this cluster
            weights = self.optimal_weights[cluster]
            
            # Process each bag in the cluster
            for i, time_series in enumerate(cluster_data['data']):
                plant = cluster_data['items'][i]['plant']
                bag = cluster_data['items'][i]['bag']
                
                # Find Feb 2025 data for this plant-bag
                feb_value = feb_2025_data.get((plant, bag), None)
                if feb_value is None:
                    continue
                
                # Generate individual forecasts
                forecasts_dict = self._generate_individual_forecasts(time_series, feb_value)
                
                # Calculate combined forecast
                combined_forecast = (
                    weights['holt_winters'] * forecasts_dict['holt_winters'] +
                    weights['random_forest'] * forecasts_dict['random_forest'] +
                    weights['trend_based'] * forecasts_dict['trend_based']
                )
                
                # Calculate confidence interval
                forecasts_array = np.array([
                    forecasts_dict['holt_winters'],
                    forecasts_dict['random_forest'],
                    forecasts_dict['trend_based']
                ])
                std_dev = np.std(forecasts_array)
                z_score = 1.96
                
                # Get March 2024 value for YoY comparison
                march_2024 = None
                march_2024_data = time_series[(time_series['Year'] == 2024) & (time_series['Month'] == 3)]
                if not march_2024_data.empty:
                    march_2024 = march_2024_data['Usage'].iloc[0]
                
                # Calculate YoY change
                yoy_change = None
                if march_2024 is not None and march_2024 > 0:
                    yoy_change = (combined_forecast - march_2024) / march_2024 * 100
                
                # Store forecast
                forecasts.append({
                    'Cluster': cluster,
                    'Plant': plant,
                    'Bag': bag,
                    'February 2025': feb_value,
                    'March 2025 Forecast': combined_forecast,
                    'Lower Bound': combined_forecast - z_score * std_dev,
                    'Upper Bound': combined_forecast + z_score * std_dev,
                    'Holt-Winters': forecasts_dict['holt_winters'],
                    'Random Forest': forecasts_dict['random_forest'],
                    'Trend-Based': forecasts_dict['trend_based'],
                    'March 2024': march_2024,
                    'YoY Change (%)': yoy_change
                })
        
        # Convert to DataFrame
        self.forecasts = pd.DataFrame(forecasts)
        return self.forecasts
    
    def prepare_results(self):
        """Prepare results for export."""
        # 1. Cluster Optimal Weights
        weights_data = []
        for cluster, weights in self.optimal_weights.items():
            weights_data.append({
                'Cluster': cluster,
                'Holt-Winters Weight': weights['holt_winters'],
                'Random Forest Weight': weights['random_forest'],
                'Trend-Based Weight': weights['trend_based'],
                'Validation MAPE (%)': weights.get('mape', 'N/A')
            })
        
        weights_df = pd.DataFrame(weights_data)
        
        # 2. Forecasts (already prepared in generate_forecasts)
        
        # 3. Cluster Summary
        if self.forecasts is not None and len(self.forecasts) > 0:
            cluster_summary = self.forecasts.groupby('Cluster').agg({
                'March 2025 Forecast': ['mean', 'std', 'count'],
                'YoY Change (%)': ['mean', 'min', 'max']
            })
            
            # Flatten multi-index columns
            cluster_summary.columns = ['_'.join(col).strip() for col in cluster_summary.columns.values]
            cluster_summary = cluster_summary.reset_index()
        else:
            cluster_summary = pd.DataFrame()
        
        self.results = {
            'weights': weights_df,
            'forecasts': self.forecasts,
            'summary': cluster_summary
        }
        
        return self.results
    
    def export_to_excel(self, filename='cluster_optimized_forecasts.xlsx'):
        """Export results to an Excel file."""
        if not self.results:
            self.prepare_results()
            
        # Create ExcelWriter for saving to a file
        with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
            # Write the forecast data
            self.results['forecasts'].to_excel(writer, sheet_name='March 2025 Forecasts', index=False)
            
            # Write the optimal weights
            self.results['weights'].to_excel(writer, sheet_name='Optimal Weights', index=False)
            
            # Write the cluster summary
            self.results['summary'].to_excel(writer, sheet_name='Cluster Summary', index=False)
            
            # Write metadata
            metadata = pd.DataFrame({
                'Metadata': ['Date Generated', 'Total Bags Forecasted', 'Total Clusters'],
                'Value': [
                    datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    len(self.results['forecasts']),
                    len(self.results['weights'])
                ]
            })
            metadata.to_excel(writer, sheet_name='Metadata', index=False)
            
            # Get the workbook and add formats
            workbook = writer.book
            
            # Format headers for all sheets
            header_format = workbook.add_format({
                'bold': True,
                'bg_color': '#D9E1F2',
                'border': 1
            })
            
            # Apply formatting to forecasts sheet
            forecast_sheet = writer.sheets['March 2025 Forecasts']
            for col_num, value in enumerate(self.results['forecasts'].columns.values):
                forecast_sheet.write(0, col_num, value, header_format)
                forecast_sheet.set_column(col_num, col_num, 15)
            
            # Add filters
            forecast_sheet.autofilter(0, 0, len(self.results['forecasts']), len(self.results['forecasts'].columns) - 1)
            
            # Apply formatting to weights sheet
            weights_sheet = writer.sheets['Optimal Weights']
            for col_num, value in enumerate(self.results['weights'].columns.values):
                weights_sheet.write(0, col_num, value, header_format)
                weights_sheet.set_column(col_num, col_num, 18)
        
        print(f"Excel file saved to {filename}")
        return filename

def main():
    # Import necessary libraries
    import pandas as pd
    from google.colab import files
    import io
    import os
    
    # Upload cluster and historical data files
    print("Please upload the cluster assignments file (Excel):")
    uploaded_clusters = files.upload()
    cluster_file = list(uploaded_clusters.keys())[0]
    
    print("\nPlease upload the historical data file (Excel with February 2025 data):")
    uploaded_historical = files.upload()
    historical_file = list(uploaded_historical.keys())[0]
    
    # Load the data
    cluster_df = pd.read_excel(io.BytesIO(uploaded_clusters[cluster_file]))
    historical_df = pd.read_excel(io.BytesIO(uploaded_historical[historical_file]))
    
    # Create optimizer
    optimizer = ClusterEnsembleOptimizer(cluster_df, historical_df)
    
    # Organize data by cluster
    optimizer.organize_by_cluster()
    
    # Optimize weights for each cluster
    optimizer.optimize_cluster_weights()
    
    # Generate forecasts using February data from historical file
    forecasts = optimizer.generate_forecasts()
    print(f"\nGenerated forecasts for {len(forecasts)} bags")
    
    # Prepare results
    results = optimizer.prepare_results()
    
    # Export to Excel file and save to disk first
    filename = 'cluster_optimized_forecasts.xlsx'
    optimizer.export_to_excel(filename)
    
    # Verify file exists before downloading
    if os.path.exists(filename):
        # Download the Excel file
        files.download(filename)
        print(f"\nResults exported to {filename}")
    else:
        print(f"\nError: Could not create file {filename}")
    
    # Display weights summary
    print("\nOptimal Weights by Cluster:")
    print(results['weights'])
    
    # Plot weight distribution
    plt.figure(figsize=(12, 6))
    
    weights_df = results['weights']
    clusters = weights_df['Cluster']
    
    bar_width = 0.25
    index = np.arange(len(clusters))
    
    plt.bar(index, weights_df['Holt-Winters Weight'], bar_width, 
            label='Holt-Winters', color='#3498db')
    plt.bar(index + bar_width, weights_df['Random Forest Weight'], bar_width,
            label='Random Forest', color='#2ecc71')
    plt.bar(index + 2*bar_width, weights_df['Trend-Based Weight'], bar_width,
            label='Trend-Based', color='#e74c3c')
    
    plt.xlabel('Cluster')
    plt.ylabel('Weight')
    plt.title('Optimal Ensemble Weights by Cluster')
    plt.xticks(index + bar_width, clusters)
    plt.legend()
    plt.tight_layout()
    plt.show()

if __name__ == "__main__":
    main()
