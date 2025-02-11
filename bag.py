import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
from statsmodels.tsa.statespace.sarimax import SARIMAX
from statsmodels.tsa.stattools import adfuller
import warnings
warnings.filterwarnings('ignore')

def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def find_best_sarima_parameters(data):
    """Find best SARIMA parameters using grid search"""
    best_aic = float('inf')
    best_params = None
    
    # Define parameter ranges
    p_values = range(0, 3)
    d_values = range(0, 2)
    q_values = range(0, 3)
    
    for p in p_values:
        for d in d_values:
            for q in q_values:
                try:
                    model = SARIMAX(data, order=(p, d, q), seasonal_order=(1, 1, 1, 12))
                    results = model.fit()
                    if results.aic < best_aic:
                        best_aic = results.aic
                        best_params = (p, d, q)
                except:
                    continue
    
    return best_params

def select_best_model(data):
    """
    Select and fit the best time series model based on data characteristics
    Returns the best model, forecast, and model description
    """
    if len(data) < 12:
        return None, None, "Forecasting not possible due to insufficient data (minimum 12 months required)"
    
    try:
        # Prepare data
        data = data.astype(float)
        
        # Initialize variables for model comparison
        best_aic = float('inf')
        best_model = None
        best_forecast = None
        best_model_name = None
        
        # 1. Try SARIMA
        try:
            best_params = find_best_sarima_parameters(data)
            if best_params:
                sarima_model = SARIMAX(data, 
                                     order=best_params, 
                                     seasonal_order=(1, 1, 1, 12))
                sarima_results = sarima_model.fit()
                if sarima_results.aic < best_aic:
                    best_aic = sarima_results.aic
                    best_model = sarima_results
                    best_forecast = sarima_results.forecast(1)[0]
                    best_model_name = f"SARIMA{best_params}x(1,1,1,12)"
        except:
            pass

        # 2. Try Holt-Winters
        try:
            hw_model = ExponentialSmoothing(data, 
                                          seasonal_periods=12, 
                                          trend='add', 
                                          seasonal='add').fit()
            if hw_model.aic < best_aic:
                best_aic = hw_model.aic
                best_model = hw_model
                best_forecast = hw_model.forecast(1)[0]
                best_model_name = "Holt-Winters Exponential Smoothing"
        except:
            pass
        
        # 3. Try simple SARIMA if others fail
        if best_model is None:
            try:
                simple_model = SARIMAX(data, order=(1, 1, 1), 
                                     seasonal_order=(1, 1, 1, 12))
                simple_results = simple_model.fit()
                best_model = simple_results
                best_forecast = simple_results.forecast(1)[0]
                best_model_name = "SARIMA(1,1,1)x(1,1,1,12)"
            except:
                pass

        if best_model is None:
            return None, None, "Could not fit any time series model to the data"

        return best_model, best_forecast, best_model_name

    except Exception as e:
        return None, None, f"Error in model fitting: {str(e)}"

def main():
    st.set_page_config("Cement Plant Bag Usage Analysis",layout="wide")
    st.title("Cement Plant Bag Usage Analysis")
    
    # Create two columns
    col1, col2 = st.columns([2, 1])
    
    with col1:
        # File uploader
        uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Remove the first column
        df = df.iloc[:, 1:]
        
        # Get unique plant names from "Cement Plant Sname" column
        unique_plants = sorted(df['Cement Plant Sname'].unique())
        
        with col1:
            # First dropdown for plant selection
            selected_plant = st.selectbox(
                'Select Cement Plant:',
                unique_plants
            )
            
            # Filter bags for selected plant
            plant_bags = df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique()
            
            # Second dropdown for bag selection
            selected_bag = st.selectbox(
                'Select Bag:',
                sorted(plant_bags)
            )
        
        # Get the row for selected plant and bag
        selected_data = df[(df['Cement Plant Sname'] == selected_plant) & 
                         (df['MAKTX'] == selected_bag)]
        
        if not selected_data.empty:
            # Get all monthly columns (excluding 'Cement Plant Sname' and 'MAKTX')
            month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
            
            # Create data for all months (for table)
            all_usage_data = []
            for month in month_columns:
                date = pd.to_datetime(month)
                usage = selected_data[month].iloc[0]
                all_usage_data.append({
                    'Date': date,
                    'Usage': usage
                })
            
            # Create DataFrame for all historical data
            all_data_df = pd.DataFrame(all_usage_data)
            all_data_df = all_data_df.sort_values('Date')
            
            # Add Month column for display
            all_data_df['Month'] = all_data_df['Date'].apply(format_date_for_display)
            
            # Filter data from Apr 2024 onwards for plotting
            apr_2024_date = pd.to_datetime('2024-04-01')
            plot_data = all_data_df[all_data_df['Date'] >= apr_2024_date].copy()
            
            # Time Series Forecasting
            with col2:
                st.subheader("Demand Forecasting")
                st.write("Results for Feb 2025:")
                
                # Prepare time series data
                ts_data = all_data_df.set_index('Date')['Usage']
                
                # Get best model and forecast
                with st.spinner('Fitting time series models...'):
                    model, forecast, model_description = select_best_model(ts_data)
                
                if forecast is not None:
                    st.write(f"**Selected Model:**")
                    st.write(model_description)
                    st.write(f"**Predicted Demand:**")
                    st.write(f"{forecast:,.2f}")
                    
                    # Add forecasted value to plot data
                    plot_data.loc[plot_data['Month'] == 'Feb 2025', 'Forecasted'] = forecast
                else:
                    st.write(model_description)
            
            with col1:
                # Create figure with custom layout
                fig = go.Figure()
                
                # Add actual usage line
                fig.add_trace(go.Scatter(
                    x=plot_data['Month'],
                    y=plot_data['Usage'],
                    name='Actual Usage',
                    line=dict(color='#1f77b4', width=3),
                    mode='lines+markers',
                    marker=dict(size=10, symbol='circle')
                ))
                
                # Add forecasted point if available
                if 'Forecasted' in plot_data.columns:
                    feb_forecast = plot_data.loc[plot_data['Month'] == 'Feb 2025', 'Forecasted'].iloc[0]
                    fig.add_trace(go.Scatter(
                        x=['Feb 2025'],
                        y=[feb_forecast],
                        name='Forecasted (Feb)',
                        mode='markers',
                        marker=dict(
                            color='red',
                            size=12,
                            symbol='star'
                        )
                    ))
                
                # Add brand rejuvenation vertical line
                fig.add_shape(
                    type="line",
                    x0="Jan 2025",
                    x1="Jan 2025",
                    y0=0,
                    y1=plot_data['Usage'].max() * 1.1,
                    line=dict(color="#FF4B4B", width=2, dash="dash"),
                )
                
                # Add annotation for brand rejuvenation
                fig.add_annotation(
                    x="Jan 2025",
                    y=plot_data['Usage'].max() * 1.15,
                    text="Brand Rejuvenation<br>(15th Jan 2025)",
                    showarrow=True,
                    arrowhead=1,
                    ax=0,
                    ay=-40,
                    font=dict(size=12, color="#FF4B4B"),
                    bgcolor="white",
                    bordercolor="#FF4B4B",
                    borderwidth=2
                )
                
                # Update layout with enhanced styling
                fig.update_layout(
                    title={
                        'text': f'Monthly Usage for {selected_bag} at {selected_plant}<br><sup>Showing data from April 2024 onwards</sup>',
                        'y':0.95,
                        'x':0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=20)
                    },
                    xaxis_title='Month',
                    yaxis_title='Usage',
                    legend_title='Type',
                    hovermode='x unified',
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    showlegend=True,
                    xaxis=dict(
                        showgrid=True,
                        gridcolor='lightgray',
                        tickangle=45
                    ),
                    yaxis=dict(
                        showgrid=True,
                        gridcolor='lightgray',
                        zeroline=True,
                        zerolinecolor='lightgray'
                    ),
                    legend=dict(
                        yanchor="top",
                        y=0.99,
                        xanchor="left",
                        x=0.01,
                        bgcolor='rgba(255, 255, 255, 0.8)'
                    )
                )
                
                # Display the graph
                st.plotly_chart(fig, use_container_width=True)
                
                # Display the complete historical data
                st.subheader("Complete Historical Data")
                # Display the data without the Date column
                display_df = all_data_df[['Month', 'Usage']]
                st.dataframe(
                    display_df.style.format({'Usage': '{:,.2f}'})
                )

if __name__ == '__main__':
    main()
