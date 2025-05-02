import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import matplotlib.ticker as mtick
from scipy.stats import pearsonr, spearmanr
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import io
import base64
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, Table, TableStyle, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas
from PIL import Image as PILImage
from io import BytesIO
import matplotlib
matplotlib.use('Agg')
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("viridis")
sns.set_context("talk")

st.set_page_config(layout="wide", page_title="Cement Consumption Model Comparison")
st.markdown("""<style>
.main {background-color: #f8f9fa;}
h1, h2, h3 {color: #2c3e50;}
.stButton>button {background-color: #3498db;color: white;}
.metric-card {background-color: white;border-radius: 5px;box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);padding: 20px;margin: 10px 0;}
.highlight {font-weight: bold;color: #27ae60;}
.month-title {text-align: center; font-size: 24px; font-weight: bold; margin-bottom: 15px; padding: 10px; background-color: #eef2f7; border-radius: 5px;}
</style>""", unsafe_allow_html=True)

st.title("Cement Bag Consumption Model Comparison Dashboard")
st.markdown("""<div style="background-color: #f0f5ff; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db;">
<h3 style="margin-top: 0;">Model Comparison</h3>
<p><strong>Model 1:</strong> Neural Network Algorithm</p>
<p><strong>Model 2:</strong> Ensemble Algorithm (Holt-Winters + Trend-Based + Random-Forest)</p>
<p>This dashboard provides a comprehensive comparison between these two prediction models for cement bag consumption against actual values for March and April.</p>
</div>""", unsafe_allow_html=True)

def calculate_metrics(actual, predicted):
    mae = mean_absolute_error(actual, predicted)
    mse = mean_squared_error(actual, predicted)
    rmse = np.sqrt(mse)
    
    # Handle zero values in actual data to avoid division by zero
    mape = np.mean(np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual))) * 100
    wmape = np.sum(np.abs(actual - predicted)) / np.sum(np.maximum(np.ones(len(actual)), actual)) * 100
    
    r2 = r2_score(actual, predicted)
    
    # Calculate correlations (handle constant arrays)
    if np.std(actual) == 0 or np.std(predicted) == 0:
        pearson_corr = 0
        spearman_corr = 0
    else:
        pearson_corr, _ = pearsonr(actual, predicted)
        spearman_corr, _ = spearmanr(actual, predicted)
    
    percent_errors = np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual)) * 100
    
    # Direction analysis
    under_predictions = np.sum(predicted < actual)
    over_predictions = np.sum(predicted > actual)
    perfect_predictions = np.sum(predicted == actual)
    
    # Error thresholds
    within_1_percent = np.sum(percent_errors <= 1)
    within_1_percent_ratio = within_1_percent / len(actual) * 100
    
    within_3_percent = np.sum(percent_errors <= 3)
    within_3_percent_ratio = within_3_percent / len(actual) * 100
    
    within_5_percent = np.sum(percent_errors <= 5)
    within_5_percent_ratio = within_5_percent / len(actual) * 100
    
    within_10_percent = np.sum(percent_errors <= 10)
    within_10_percent_ratio = within_10_percent / len(actual) * 100
    
    # Total values
    total_actual = np.sum(actual)
    total_predicted = np.sum(predicted)
    abs_total_deviation = abs(total_actual - total_predicted)
    total_deviation_percent = (abs_total_deviation / total_actual) * 100 if total_actual > 0 else 0
    
    # Bias metrics
    bias = np.mean(predicted - actual)
    bias_percent = (bias / np.mean(actual)) * 100 if np.mean(actual) > 0 else 0
    
    # Tracking signal
    sum_errors = np.sum(predicted - actual)
    mad = np.mean(np.abs(predicted - actual))
    tracking_signal = sum_errors / mad if mad > 0 else 0
    
    return {
        'MAE': mae,
        'MSE': mse,
        'RMSE': rmse,
        'MAPE': mape,
        'WMAPE': wmape,
        'R²': r2,
        'Pearson Correlation': pearson_corr,
        'Spearman Correlation': spearman_corr,
        'Under Predictions': under_predictions,
        'Over Predictions': over_predictions,
        'Perfect Predictions': perfect_predictions,
        'Within 1% Error (%)': within_1_percent_ratio,
        'Within 3% Error (%)': within_3_percent_ratio,
        'Within 5% Error (%)': within_5_percent_ratio,
        'Within 10% Error (%)': within_10_percent_ratio,
        'Total Deviation (%)': total_deviation_percent,
        'Bias': bias,
        'Bias (%)': bias_percent,
        'Tracking Signal': tracking_signal,
        'Percent Errors': percent_errors
    }

def process_month_data(df, month_prefix):
    """Process data for a specific month (Mar or Apr)"""
    
    actual_col = f"{month_prefix}-Actual"
    pred1_col = f"{month_prefix} Pred1"
    pred2_col = f"{month_prefix} Pred2"
    
    # Create error columns
    df[f'Error_Model1_{month_prefix}'] = df[actual_col] - df[pred1_col]
    df[f'Error_Model2_{month_prefix}'] = df[actual_col] - df[pred2_col]
    
    # Calculate error percentages
    df[f'Error_Percent_Model1_{month_prefix}'] = np.abs(df[f'Error_Model1_{month_prefix}'] / np.maximum(np.ones(len(df)), df[actual_col])) * 100
    df[f'Error_Percent_Model2_{month_prefix}'] = np.abs(df[f'Error_Model2_{month_prefix}'] / np.maximum(np.ones(len(df)), df[actual_col])) * 100
    
    # Calculate metrics
    metrics_model1 = calculate_metrics(df[actual_col], df[pred1_col])
    metrics_model2 = calculate_metrics(df[actual_col], df[pred2_col])
    
    return df, metrics_model1, metrics_model2

def create_month_dashboard(df, month_prefix, metrics_model1, metrics_model2, container):
    """Create dashboard visualizations for a specific month"""
    
    with container:
        st.markdown(f"<div class='month-title'>{month_prefix}ch Analysis</div>", unsafe_allow_html=True)
        
        # Display metrics
        col1, col2 = st.columns(2)
        with col1:
            st.subheader("Model 1 Performance Metrics")
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            for metric, value in metrics_model1.items():
                if metric != 'Percent Errors':
                    if isinstance(value, (int, float)):
                        st.metric(metric, f"{value:.4f}" if value < 100 else f"{value:.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.subheader("Model 2 Performance Metrics")
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            for metric, value in metrics_model2.items():
                if metric != 'Percent Errors':
                    if isinstance(value, (int, float)):
                        st.metric(metric, f"{value:.4f}" if value < 100 else f"{value:.2f}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        # Model Comparison Summary
        st.subheader("Model Comparison Summary")
        comparison_df = pd.DataFrame({
            'Metric': [k for k in metrics_model1.keys() if k != 'Percent Errors'],
            'Model 1': [metrics_model1[k] for k in metrics_model1.keys() if k != 'Percent Errors'],
            'Model 2': [metrics_model2[k] for k in metrics_model1.keys() if k != 'Percent Errors']
        })
        
        better_model = []
        for metric in comparison_df['Metric']:
            if metric in ['R²', 'Pearson Correlation', 'Spearman Correlation', 'Perfect Predictions', 
                          'Within 1% Error (%)', 'Within 3% Error (%)', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                # Higher is better
                if metrics_model1[metric] > metrics_model2[metric]:
                    better_model.append("Model 1")
                elif metrics_model2[metric] > metrics_model1[metric]:
                    better_model.append("Model 2")
                else:
                    better_model.append("Equal")
            else:
                # Lower is better
                if metrics_model1[metric] < metrics_model2[metric]:
                    better_model.append("Model 1")
                elif metrics_model2[metric] < metrics_model1[metric]:
                    better_model.append("Model 2")
                else:
                    better_model.append("Equal")
        
        comparison_df['Better Model'] = better_model
        st.dataframe(comparison_df.style.apply(
            lambda x: ['background-color: #d4f1dd' if v == "Model 1" 
                      else 'background-color: #d1e7f0' if v == "Model 2" 
                      else '' for v in x], 
            subset=['Better Model']))
        
        # Winner determination
        model1_wins = sum(1 for model in better_model if model == "Model 1")
        model2_wins = sum(1 for model in better_model if model == "Model 2")
        equal_metrics = sum(1 for model in better_model if model == "Equal")
        
        winner = "Model 1" if model1_wins > model2_wins else "Model 2" if model2_wins > model1_wins else "Both models perform equally"
        
        st.markdown(f"""<div class="metric-card">
        <h3>Overall Winner: <span class="highlight">{winner}</span></h3>
        <p>Model 1 better in {model1_wins} metrics</p>
        <p>Model 2 better in {model2_wins} metrics</p>
        <p>Equal performance in {equal_metrics} metrics</p>
        </div>""", unsafe_allow_html=True)
        
        # Visualizations
        st.header("Visualizations")
        
        # Actual vs Predicted Values
        st.subheader("Actual vs Predicted Values")
        actual_col = f"{month_prefix}-Actual"
        pred1_col = f"{month_prefix} Pred1"
        pred2_col = f"{month_prefix} Pred2"
        
        df_melted = pd.melt(
            df, 
            id_vars=['Bag Plus Plant'], 
            value_vars=[actual_col, pred1_col, pred2_col],
            var_name='Measurement Type', 
            value_name='Value'
        )
        
        fig = px.bar(
            df_melted, 
            x='Bag Plus Plant', 
            y='Value', 
            color='Measurement Type', 
            barmode='group',
            color_discrete_map={
                actual_col: '#2ecc71', 
                pred1_col: '#3498db', 
                pred2_col: '#9b59b6'
            },
            title=f'Comparison of Actual vs Predicted Consumption by Cement Bag Type - {month_prefix}ch'
        )
        
        fig.update_layout(
            xaxis_title='Cement Bag Type', 
            yaxis_title='Consumption',
            legend_title='Data Type', 
            template='plotly_white'
        )
        
        if len(df) > 5:
            fig.update_layout(xaxis_tickangle=-45)
            
        st.plotly_chart(fig, use_container_width=True)
        
        # Error Analysis
        st.subheader("Error Analysis")
        error_fig = make_subplots(
            rows=1, 
            cols=2, 
            subplot_titles=(f"Model 1 Error Distribution - {month_prefix}ch", f"Model 2 Error Distribution - {month_prefix}ch")
        )
        
        error_fig.add_trace(
            go.Histogram(
                x=df[f'Error_Model1_{month_prefix}'], 
                name='Model 1 Error', 
                marker_color='#3498db'
            ),
            row=1, col=1
        )
        
        error_fig.add_trace(
            go.Histogram(
                x=df[f'Error_Model2_{month_prefix}'], 
                name='Model 2 Error', 
                marker_color='#9b59b6'
            ),
            row=1, col=2
        )
        
        error_fig.update_layout(
            height=500, 
            title_text=f"Error Distribution Comparison - {month_prefix}ch",
            template='plotly_white'
        )
        
        st.plotly_chart(error_fig, use_container_width=True)
        
        # Scatter Plots
        scatter_fig = make_subplots(
            rows=1, 
            cols=2, 
            subplot_titles=(f"Model 1: Actual vs Predicted - {month_prefix}ch", f"Model 2: Actual vs Predicted - {month_prefix}ch"),
            specs=[[{"type": "scatter"}, {"type": "scatter"}]]
        )
        
        max_val = max(df[actual_col].max(), df[pred1_col].max(), df[pred2_col].max())
        min_val = min(df[actual_col].min(), df[pred1_col].min(), df[pred2_col].min())
        
        for col in [1, 2]:
            scatter_fig.add_trace(
                go.Scatter(
                    x=[min_val, max_val], 
                    y=[min_val, max_val], 
                    mode='lines', 
                    name='Perfect Prediction',
                    line=dict(color='rgba(0,0,0,0.5)', dash='dash'), 
                    showlegend=col==1
                ),
                row=1, col=col
            )
        
        scatter_fig.add_trace(
            go.Scatter(
                x=df[actual_col], 
                y=df[pred1_col], 
                mode='markers', 
                name='Model 1',
                marker=dict(color='#3498db', size=10)
            ),
            row=1, col=1
        )
        
        scatter_fig.add_trace(
            go.Scatter(
                x=df[actual_col], 
                y=df[pred2_col], 
                mode='markers', 
                name='Model 2',
                marker=dict(color='#9b59b6', size=10)
            ),
            row=1, col=2
        )
        
        scatter_fig.update_layout(
            height=500, 
            title_text=f"Actual vs Predicted Scatter Plots - {month_prefix}ch",
            xaxis_title="Actual Values", 
            yaxis_title="Predicted Values",
            xaxis2_title="Actual Values", 
            yaxis2_title="Predicted Values",
            template='plotly_white'
        )
        
        st.plotly_chart(scatter_fig, use_container_width=True)
        
        # Percentage Error by Cement Bag Type
        st.subheader("Percentage Error by Cement Bag Type")
        percent_error_df = pd.DataFrame({
            'Bag Plus Plant': df['Bag Plus Plant'],
            'Model 1 Error (%)': df[f'Error_Percent_Model1_{month_prefix}'],
            'Model 2 Error (%)': df[f'Error_Percent_Model2_{month_prefix}']
        })
        
        percent_error_melted = pd.melt(
            percent_error_df, 
            id_vars=['Bag Plus Plant'],
            value_vars=['Model 1 Error (%)', 'Model 2 Error (%)'],
            var_name='Model', 
            value_name='Percentage Error'
        )
        
        percent_fig = px.bar(
            percent_error_melted, 
            x='Bag Plus Plant', 
            y='Percentage Error', 
            color='Model',
            barmode='group', 
            title=f'Percentage Error Comparison by Cement Bag Type - {month_prefix}ch',
            color_discrete_map={
                'Model 1 Error (%)': '#3498db', 
                'Model 2 Error (%)': '#9b59b6'
            }
        ) 
        
        percent_fig.update_layout(
            xaxis_title='Cement Bag Type', 
            yaxis_title='Percentage Error (%)',
            legend_title='Model', 
            template='plotly_white'
        )
        
        if len(df) > 5:
            percent_fig.update_layout(xaxis_tickangle=-45)
            
        st.plotly_chart(percent_fig, use_container_width=True)
        
        # Radar Chart: Model Performance Comparison
        st.subheader("Radar Chart: Model Performance Comparison")
        
        radar_metrics = ['MAE', 'RMSE', 'MAPE', 'R²', 'Within 5% Error (%)', 'Within 10% Error (%)']
        radar_df = pd.DataFrame({
            'Metric': radar_metrics,
            'Model 1': [metrics_model1[m] for m in radar_metrics],
            'Model 2': [metrics_model2[m] for m in radar_metrics]
        })
        
        # Normalize metrics for radar chart
        for metric in radar_metrics:
            if metric in ['R²', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                # Higher is better
                max_val = max(
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0]
                )
                if max_val != 0:
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val
            else:
                # Lower is better
                max_val = max(
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0]
                )
                if max_val != 0:
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val)
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val)
        
        radar_fig = go.Figure()
        
        radar_fig.add_trace(go.Scatterpolar(
            r=radar_df['Model 1'].values,
            theta=radar_df['Metric'].values,
            fill='toself',
            name='Model 1',
            line_color='#3498db'
        ))
        
        radar_fig.add_trace(go.Scatterpolar(
            r=radar_df['Model 2'].values,
            theta=radar_df['Metric'].values,
            fill='toself',
            name='Model 2',
            line_color='#9b59b6'
        ))
        
        radar_fig.update_layout(
            polar=dict(
                radialaxis=dict(
                    visible=True,
                    range=[0, 1]
                )
            ),
            showlegend=True,
            title=f'Model Performance Radar Chart - {month_prefix}ch (Higher is Better for All Metrics)',
            template='plotly_white'
        )
        
        st.plotly_chart(radar_fig, use_container_width=True)
        
        # Error Trend Analysis for High Volume Products
        st.subheader("Error Trend Analysis for High Volume Products")
        
        top_products = min(5, len(df))
        top_df = df.sort_values(actual_col, ascending=False).head(top_products)
        
        trend_fig = make_subplots(specs=[[{"secondary_y": True}]])
        
        trend_fig.add_trace(
            go.Bar(
                x=top_df['Bag Plus Plant'], 
                y=top_df[actual_col], 
                name='Actual Consumption',
                marker_color='rgba(46, 204, 113, 0.7)'
            ),
            secondary_y=False,
        )
        
        trend_fig.add_trace(
            go.Scatter(
                x=top_df['Bag Plus Plant'], 
                y=top_df[f'Error_Percent_Model1_{month_prefix}'],
                mode='lines+markers', 
                name='Neural Network Error (%)',
                line=dict(color='#3498db', width=2)
            ),
            secondary_y=True,
        )
        
        trend_fig.add_trace(
            go.Scatter(
                x=top_df['Bag Plus Plant'], 
                y=top_df[f'Error_Percent_Model2_{month_prefix}'], 
                mode='lines+markers', 
                name='Ensemble Error (%)',
                line=dict(color='#9b59b6', width=2)
            ),
            secondary_y=True,
        )
        
        trend_fig.update_layout(
            title_text=f"High Volume Products: Actual Consumption vs Error Percentage - {month_prefix}ch",
            template='plotly_white',
            barmode='group'
        )
        
        trend_fig.update_yaxes(title_text="Actual Consumption", secondary_y=False)
        trend_fig.update_yaxes(title_text="Error Percentage (%)", secondary_y=True)
        
        st.plotly_chart(trend_fig, use_container_width=True)

# Upload files
st.subheader("Upload Excel Files")

# Layout for file upload
col1, col2 = st.columns(2)

with col1:
    st.markdown("<div class='month-title'>March Data</div>", unsafe_allow_html=True)
    march_file = st.file_uploader("Upload March Excel file", type=["xlsx", "xls"], key="march_file")

with col2:
    st.markdown("<div class='month-title'>April Data</div>", unsafe_allow_html=True)
    april_file = st.file_uploader("Upload April Excel file", type=["xlsx", "xls"], key="april_file")

# Process files if uploaded
if march_file and april_file:
    try:
        # Read both files
        df_march = pd.read_excel(march_file)
        df_april = pd.read_excel(april_file)
        
        # Check required columns for March data
        march_required_columns = ["Bag Plus Plant", "Mar-Actual", "Mar Pred1", "Mar Pred2"]
        april_required_columns = ["Bag Plus Plant", "Apr-Actual", "Apr Pred1", "Apr Pred2"]
        
        if all(col in df_march.columns for col in march_required_columns) and all(col in df_april.columns for col in april_required_columns):
            # Process data for both months
            df_march, metrics_model1_march, metrics_model2_march = process_month_data(df_march, "Mar")
            df_april, metrics_model1_april, metrics_model2_april = process_month_data(df_april, "Apr")
            
            # Show raw data in expandable sections
            st.subheader("Raw Data Preview")
            
            raw_col1, raw_col2 = st.columns(2)
            
            with raw_col1:
                with st.expander("March Raw Data"):
                    st.dataframe(df_march)
            
            with raw_col2:
                with st.expander("April Raw Data"):
                    st.dataframe(df_april)
            
            # Create dashboard for each month
            dashboard_col1, dashboard_col2 = st.columns(2)
            
            create_month_dashboard(df_march, "Mar", metrics_model1_march, metrics_model2_march, dashboard_col1)
            create_month_dashboard(df_april, "Apr", metrics_model1_april, metrics_model2_april, dashboard_col2)
            
            # Download reports
            st.header("Download Reports")
            
            # March report
            march_output = io.BytesIO()
            with pd.ExcelWriter(march_output, engine='xlsxwriter') as writer:
                df_march.to_excel(writer, sheet_name='March Data', index=False)
                
                comparison_df_march = pd.DataFrame({
                    'Metric': [k for k in metrics_model1_march.keys() if k != 'Percent Errors'],
                    'Model 1': [metrics_model1_march[k] for k in metrics_model1_march.keys() if k != 'Percent Errors'],
                    'Model 2': [metrics_model2_march[k] for k in metrics_model1_march.keys() if k != 'Percent Errors']
                })
                
                comparison_df_march.to_excel(writer, sheet_name='March Metrics', index=False)
            
            march_excel_data = march_output.getvalue()
            
            # April report
            april_output = io.BytesIO()
            with pd.ExcelWriter(april_output, engine='xlsxwriter') as writer:
                df_april.to_excel(writer, sheet_name='April Data', index=False)
                
                comparison_df_april = pd.DataFrame({
                    'Metric': [k for k in metrics_model1_april.keys() if k != 'Percent Errors'],
                    'Model 1': [metrics_model1_april[k] for k in metrics_model1_april.keys() if k != 'Percent Errors'],
                    'Model 2': [metrics_model2_april[k] for k in metrics_model1_april.keys() if k != 'Percent Errors']
                })
                
                comparison_df_april.to_excel(writer, sheet_name='April Metrics', index=False)
            
            april_excel_data = april_output.getvalue()
            
            # Combined report
            combined_output = io.BytesIO()
            with pd.ExcelWriter(combined_output, engine='xlsxwriter') as writer:
                df_march.to_excel(writer, sheet_name='March Data', index=False)
                df_april.to_excel(writer, sheet_name='April Data', index=False)
                
                comparison_df_march.to_excel(writer, sheet_name='March Metrics', index=False)
                comparison_df_april.to_excel(writer, sheet_name='April Metrics', index=False)
            
            combined_excel_data = combined_output.getvalue()
            
            # Download buttons
            download_col1, download_col2, download_col3 = st.columns(3)
            
            with download_col1:
                st.download_button(
                    label="Download March Analysis",
                    data=march_excel_data,
                    file_name="cement_model_comparison_march.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with download_col2:
                st.download_button(
                    label="Download April Analysis",
                    data=april_excel_data,
                    file_name="cement_model_comparison_april.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            
            with download_col3:
                st.download_button(
                    label="Download Combined Analysis",
                    data=combined_excel_data,
                    file_name="cement_model_comparison_combined.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
        else:
            missing_march = [col for col in march_required_columns if col not in df_march.columns]
            missing_april = [col for col in april_required_columns if col not in df_april.columns]
            
            if missing_march:
                st.error(f"Required columns not found in March file: {', '.join(missing_march)}")
            if missing_april:
                st.error(f"Required columns not found in April file: {', '.join(missing_april)}")
    except Exception as e:
        st.error(f"Error processing the file: {str(e)}")
else:
    st.info("Please upload an Excel file with the following columns: 'Bag Plus Plant', 'Mar-Actual', 'Mar Pred1', 'Mar Pred2'")
    sample_df = pd.DataFrame({'Bag Plus Plant': ['Cement Type A - Plant 1', 'Cement Type B - Plant 2', 'Cement Type C - Plant 1'],'Mar-Actual': [1500, 2000, 1200],'Mar Pred1': [1450, 2100, 1250],'Mar Pred2': [1530, 1950, 1180]})
    st.write("Sample data structure:")
    st.dataframe(sample_df)
st.markdown("""<div style="text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; border-radius: 5px;"><p style="color: #7f8c8d;">Cement Consumption Model Comparison Dashboard</p></div>""", unsafe_allow_html=True)
