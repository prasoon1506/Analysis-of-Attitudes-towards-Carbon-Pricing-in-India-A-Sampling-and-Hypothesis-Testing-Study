import pandas as pd
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.seasonal import seasonal_decompose
from sklearn.ensemble import RandomForestRegressor
import matplotlib.pyplot as plt
from datetime import datetime
import ipywidgets as widgets
from IPython.display import display, clear_output
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
    def _extrapolate_april(self, apr_partial_data):
        days_in_apr = 30
        current_days = 20
        daily_rate = apr_partial_data / current_days
        return daily_rate * days_in_apr
    def _get_recent_trend(self, window=6):
        recent_data = self.data.tail(window)
        return (recent_data['Usage'].iloc[-1] - recent_data['Usage'].iloc[0]) / window
    def generate_forecast(self, apr_partial_data):
        apr_full = self._extrapolate_april(apr_partial_data)
        seasonal_indices = self._calculate_seasonal_indices()
        may_seasonal_factor = 1.0 if isinstance(seasonal_indices, float) else seasonal_indices.get(5, 1.0)
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
            may_features = pd.DataFrame({'Month': [3],'Year': [2025],'Lag1': [apr_full],'Lag12': [self.data['Usage'].iloc[-12]],'Trend': [len(self.data)]})
            rf_forecast = rf_model.predict(may_features)[0]
        except:
            rf_forecast = self.data['Usage'].mean() * may_seasonal_factor
        trend = self._get_recent_trend()
        trend_forecast = apr_full + trend
        weights = {'holt_winters': 0.4,'random_forest': 0.4,'trend_based': 0.2}
        combined_forecast = (weights['holt_winters'] * hw_forecast +weights['random_forest'] * rf_forecast +weights['trend_based'] * trend_forecast)
        forecasts = np.array([hw_forecast, rf_forecast, trend_forecast])
        std_dev = np.std(forecasts)
        z_score = 1.96
        return {'point_forecast': combined_forecast,'lower_bound': combined_forecast - z_score * std_dev,'upper_bound': combined_forecast + z_score * std_dev,'individual_forecasts': {'holt_winters': hw_forecast,'random_forest': rf_forecast,'trend_based': trend_forecast},'april_extrapolated': apr_full,'seasonal_factor': may_seasonal_factor}
def plot_forecast(usage_df, forecast_results):
    plt.figure(figsize=(12, 6))
    plt.plot(usage_df['Date'], usage_df['Usage'],label='Historical', color='#2E86C1', linewidth=2)
    apr_date = pd.to_datetime('2025-04-01')
    plt.scatter(apr_date, forecast_results['april_extrapolated'],color='orange', s=100, marker='D', label='Apr Projection')
    may_date = pd.to_datetime('2025-05-01')
    plt.scatter(may_date, forecast_results['point_forecast'],color='red', s=150, marker='*', label='May Forecast')
    plt.vlines(may_date,forecast_results['lower_bound'],forecast_results['upper_bound'],color='red', alpha=0.2, linewidth=10, label='95% Confidence')
    plt.title('Historical Data and Forecast')
    plt.xlabel('Date')
    plt.ylabel('Usage')
    plt.legend()
    plt.grid(True, alpha=0.3)
    plt.tight_layout()
    plt.show()
def display_metrics(forecast_results, prev_may):
    print("\nForecast Metrics:")
    print(f"May 2025 Forecast: {forecast_results['point_forecast']:,.0f}")
    print(f"April 2025 (Projected): {forecast_results['april_extrapolated']:,.0f}")
    yoy_change = ((forecast_results['point_forecast'] - prev_may) / prev_may * 100)
    print(f"Year-over-Year Change: {yoy_change:,.1f}%")
    print("\nForecasting Model Details:")
    for model, value in forecast_results['individual_forecasts'].items():
        print(f"{model.replace('_', ' ').title()}: {value:,.0f}")
    print(f"\nConfidence Interval: {forecast_results['lower_bound']:,.0f} - {forecast_results['upper_bound']:,.0f}")
def process_data(df, selected_plant, selected_bag):
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
    return usage_df.sort_values('Date')
def create_forecast(file_path):
    df = pd.read_excel("/content/Bag data for May planning NEW.xlsx")
    df = df.iloc[:, 1:]
    plants = sorted(df['Cement Plant Sname'].unique())
    plant_dropdown = widgets.Dropdown(options=plants,description='Plant:',style={'description_width': 'initial'})
    def update_bag_dropdown(*args):
        plant_bags = sorted(df[df['Cement Plant Sname'] == plant_dropdown.value]['MAKTX'].unique())
        bag_dropdown.options = plant_bags
    initial_plant_bags = sorted(df[df['Cement Plant Sname'] == plants[0]]['MAKTX'].unique())
    bag_dropdown = widgets.Dropdown(options=initial_plant_bags,description='Bag:',style={'description_width': 'initial'})
    forecast_button = widgets.Button(description='Generate Forecast')
    def on_forecast_button_clicked(b):
        clear_output(wait=True)
        display(plant_dropdown, bag_dropdown, forecast_button)
        usage_df = process_data(df, plant_dropdown.value, bag_dropdown.value)
        apr_2025_data = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2025-04']['Usage'].iloc[0]
        forecaster = BagDemandForecaster(usage_df)
        forecast_results = forecaster.generate_forecast(apr_2025_data)
        prev_may = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2024-05']['Usage'].iloc[0]
        plot_forecast(usage_df, forecast_results)
        display_metrics(forecast_results, prev_may)
    plant_dropdown.observe(update_bag_dropdown, 'value')
    forecast_button.on_click(on_forecast_button_clicked)
    display(plant_dropdown, bag_dropdown, forecast_button)
create_forecast('your_excel_file.xlsx')
def generate_all_forecasts(df):
    all_forecasts = []
    total_combinations = sum(len(df[df['Cement Plant Sname'] == plant]['MAKTX'].unique()) for plant in df['Cement Plant Sname'].unique())
    current_combination = 0
    for plant in sorted(df['Cement Plant Sname'].unique()):
        for bag in sorted(df[df['Cement Plant Sname'] == plant]['MAKTX'].unique()):
            current_combination += 1
            print(f"Processing combination {current_combination}/{total_combinations}: {plant} - {bag}")
            try:
                usage_df = process_data(df, plant, bag)
                apr_2025_data = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2025-04']['Usage'].iloc[0]
                feb_data = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2025-02']['Usage'].iloc[0]
                mar_data = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2025-03']['Usage'].iloc[0]
                forecaster = BagDemandForecaster(usage_df)
                forecast_results = forecaster.generate_forecast(apr_2025_data)
                prev_may = usage_df[usage_df['Date'].dt.strftime('%Y-%m') == '2024-05']['Usage'].iloc[0]
                yoy_change = ((forecast_results['point_forecast'] - prev_may) / prev_may * 100)
                all_forecasts.append({'Plant': plant,'Bag': bag,'May 2025 Forecast': forecast_results['point_forecast'],'April 2025 Projected': forecast_results['april_extrapolated'],'Year-over-Year Change (%)': yoy_change,'Holt-Winters Forecast': forecast_results['individual_forecasts']['holt_winters'],'Random Forest Forecast': forecast_results['individual_forecasts']['random_forest'],'Trend-Based Forecast': forecast_results['individual_forecasts']['trend_based'],'Confidence Interval Lower': forecast_results['lower_bound'],'Confidence Interval Upper': forecast_results['upper_bound'],'Previous May (2025)': prev_may,'February 2025': feb_data,'March 2025': mar_data})
            except Exception as e:
                print(f"Error processing {plant} - {bag}: {str(e)}")
                continue
    return pd.DataFrame(all_forecasts)
def export_forecasts_to_excel(df, output_filename='bag_demand_forecasts.xlsx'):
    print("Generating forecasts for all plant-bag combinations...")
    forecasts_df = generate_all_forecasts(df)
    writer = pd.ExcelWriter(output_filename, engine='xlsxwriter')
    forecasts_df.to_excel(writer, sheet_name='Forecasts', index=False)
    workbook = writer.book
    worksheet = writer.sheets['Forecasts']
    header_format = workbook.add_format({'bold': True,'text_wrap': True,'valign': 'top','bg_color': '#D9E1F2','border': 1})
    number_format = workbook.add_format({'num_format': '#,##0','border': 1})
    percent_format = workbook.add_format({'num_format': '0.0%','border': 1})
    for col_num, value in enumerate(forecasts_df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        max_length = max(forecasts_df[value].astype(str).apply(len).max(),len(value))
        worksheet.set_column(col_num, col_num, max_length + 2)
        if 'Forecast' in value or 'Previous' in value or 'Confidence' in value:
            worksheet.set_column(col_num, col_num, None, number_format)
        elif 'Change' in value:
            worksheet.set_column(col_num, col_num, None, percent_format)
    summary_data = {'Metric': ['Total Plants','Total Bags','Average May 2025 Forecast','Total May 2025 Forecast','Average Year-over-Year Change'],
        'Value': [len(forecasts_df['Plant'].unique()),len(forecasts_df),forecasts_df['May 2025 Forecast'].mean(),forecasts_df['May 2025 Forecast'].sum(),forecasts_df['Year-over-Year Change (%)'].mean()]}
    summary_df = pd.DataFrame(summary_data)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)
    summary_sheet = writer.sheets['Summary']
    summary_sheet.set_column('A:A', 30)
    summary_sheet.set_column('B:B', 15)
    writer.close()
    print(f"\nForecasts have been exported to {output_filename}")
    return output_filename
def create_forecast(file_path):
    df = pd.read_excel("/content/Bag data for May planning NEW.xlsx")
    df = df.iloc[:, 1:]  # Remove first column if it's an index
    export_button = widgets.Button(description='Export All Forecasts')
    plants = sorted(df['Cement Plant Sname'].unique())
    plant_dropdown = widgets.Dropdown(options=plants,description='Plant:',style={'description_width': 'initial'})
    def update_bag_dropdown(*args):
        plant_bags = sorted(df[df['Cement Plant Sname'] == plant_dropdown.value]['MAKTX'].unique())
        bag_dropdown.options = plant_bags
    initial_plant_bags = sorted(df[df['Cement Plant Sname'] == plants[0]]['MAKTX'].unique())
    bag_dropdown = widgets.Dropdown(
        options=initial_plant_bags,
        description='Bag:',
        style={'description_width': 'initial'}
    )
    
    # Create forecast button
    forecast_button = widgets.Button(description='Generate Forecast')
    
    def on_forecast_button_clicked(b):
        clear_output(wait=True)
        display(plant_dropdown, bag_dropdown, forecast_button, export_button)
        
        usage_df = process_data(df, plant_dropdown.value, bag_dropdown.value)
        
        # Get February 2025 data
        apr_2025_data = usage_df[
            usage_df['Date'].dt.strftime('%Y-%m') == '2025-04'
        ]['Usage'].iloc[0]
        
        # Generate forecast
        forecaster = BagDemandForecaster(usage_df)
        forecast_results = forecaster.generate_forecast(apr_2025_data)
        
        # Get previous March data for YoY comparison
        prev_may = usage_df[
            usage_df['Date'].dt.strftime('%Y-%m') == '2024-05'
        ]['Usage'].iloc[0]
        
        # Display results
        plot_forecast(usage_df, forecast_results)
        display_metrics(forecast_results, prev_may)
    
    def on_export_button_clicked(b):
        clear_output(wait=True)
        display(plant_dropdown, bag_dropdown, forecast_button, export_button)
        
        print("Starting export process...")
        output_filename = export_forecasts_to_excel(df)
        
        # For Google Colab, add download link
        try:
            from google.colab import files
            files.download(output_filename)
        except:
            print("File saved locally. If you're running this in Colab, the download should start automatically.")
    
    # Connect the callbacks
    plant_dropdown.observe(update_bag_dropdown, 'value')
    forecast_button.on_click(on_forecast_button_clicked)
    export_button.on_click(on_export_button_clicked)
    
    # Display the widgets
    display(plant_dropdown, bag_dropdown, forecast_button, export_button)

create_forecast('your_excel_file.xlsx')
