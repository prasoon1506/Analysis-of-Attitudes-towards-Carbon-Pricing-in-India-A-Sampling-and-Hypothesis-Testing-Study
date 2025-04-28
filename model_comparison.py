import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import matplotlib.ticker as mtick
from scipy.stats import pearsonr
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import io
plt.style.use('seaborn-v0_8-whitegrid')
sns.set_palette("viridis")
sns.set_context("talk")
st.set_page_config(layout="wide", page_title="Cement Consumption Model Comparison")
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    h1, h2, h3 {
        color: #2c3e50;
    }
    .stButton>button {
        background-color: #3498db;
        color: white;
    }
    .metric-card {
        background-color: white;
        border-radius: 5px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        padding: 20px;
        margin: 10px 0;
    }
    .highlight {
        font-weight: bold;
        color: #27ae60;
    }
</style>
""", unsafe_allow_html=True)
st.title("Cement Bag Consumption Model Comparison Dashboard")
st.markdown("This dashboard provides a comprehensive comparison between two prediction models for cement bag consumption against actual values for March.")
def calculate_metrics(actual, predicted):
    mae = mean_absolute_error(actual, predicted)
    mse = mean_squared_error(actual, predicted)
    rmse = np.sqrt(mse)
    mape = np.mean(np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual))) * 100
    r2 = r2_score(actual, predicted)
    corr, _ = pearsonr(actual, predicted)
    percent_errors = np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual)) * 100
    under_predictions = np.sum(predicted < actual)
    over_predictions = np.sum(predicted > actual)
    perfect_predictions = np.sum(predicted == actual)
    within_5_percent = np.sum(percent_errors <= 5)
    within_5_percent_ratio = within_5_percent / len(actual) * 100
    
    # Calculate percentage of predictions within 10% error
    within_10_percent = np.sum(percent_errors <= 10)
    within_10_percent_ratio = within_10_percent / len(actual) * 100
    
    return {
        'MAE': mae,
        'MSE': mse,
        'RMSE': rmse,
        'MAPE': mape,
        'R²': r2,
        'Correlation': corr,
        'Under Predictions': under_predictions,
        'Over Predictions': over_predictions,
        'Perfect Predictions': perfect_predictions,
        'Within 5% Error (%)': within_5_percent_ratio,
        'Within 10% Error (%)': within_10_percent_ratio,
        'Percent Errors': percent_errors
    }

# File uploader
st.subheader("Upload Excel File")
uploaded_file = st.file_uploader("Upload your Excel file with cement bag data", type=["xlsx", "xls"])

if uploaded_file:
    # Read the Excel file
    try:
        df = pd.read_excel(uploaded_file)
        
        # Display the raw data
        with st.expander("Raw Data Preview"):
            st.dataframe(df)
        
        # Check if all required columns are present
        required_columns = ["Bag Plus Plant", "Mar-Actual", "Mar Pred1", "Mar Pred2"]
        if all(col in df.columns for col in required_columns):
            # Create new columns for errors and percentage errors
            df['Error_Model1'] = df['Mar-Actual'] - df['Mar Pred1']
            df['Error_Model2'] = df['Mar-Actual'] - df['Mar Pred2']
            
            df['Error_Percent_Model1'] = np.abs(df['Error_Model1'] / np.maximum(np.ones(len(df)), df['Mar-Actual'])) * 100
            df['Error_Percent_Model2'] = np.abs(df['Error_Model2'] / np.maximum(np.ones(len(df)), df['Mar-Actual'])) * 100
            
            # Calculate metrics
            metrics_model1 = calculate_metrics(df['Mar-Actual'], df['Mar Pred1'])
            metrics_model2 = calculate_metrics(df['Mar-Actual'], df['Mar Pred2'])
            
            # Create columns for dashboard layout
            col1, col2 = st.columns(2)
            
            # Summary statistics
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
            
            # Create a comparison dataframe
            comparison_df = pd.DataFrame({
                'Metric': [k for k in metrics_model1.keys() if k != 'Percent Errors'],
                'Model 1': [metrics_model1[k] for k in metrics_model1.keys() if k != 'Percent Errors'],
                'Model 2': [metrics_model2[k] for k in metrics_model1.keys() if k != 'Percent Errors']
            })
            
            # Add a "Better Model" column
            better_model = []
            for metric in comparison_df['Metric']:
                if metric in ['R²', 'Correlation', 'Perfect Predictions', 'Within 5% Error (%)', 'Within 10% Error (%)']:
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
            
            # Display comparison table
            st.dataframe(comparison_df.style.apply(lambda x: ['background-color: #d4f1dd' if v == "Model 1" 
                                                    else 'background-color: #d1e7f0' if v == "Model 2" 
                                                    else '' for v in x], subset=['Better Model']))
            
            # Overall winner
            model1_wins = sum(1 for model in better_model if model == "Model 1")
            model2_wins = sum(1 for model in better_model if model == "Model 2")
            equal_metrics = sum(1 for model in better_model if model == "Equal")
            
            winner = "Model 1" if model1_wins > model2_wins else "Model 2" if model2_wins > model1_wins else "Both models perform equally"
            
            st.markdown(f"""
            <div class="metric-card">
                <h3>Overall Winner: <span class="highlight">{winner}</span></h3>
                <p>Model 1 better in {model1_wins} metrics</p>
                <p>Model 2 better in {model2_wins} metrics</p>
                <p>Equal performance in {equal_metrics} metrics</p>
            </div>
            """, unsafe_allow_html=True)
            
            # Visualization Section
            st.header("Visualizations")
            
            # 1. Actual vs Predicted Comparison
            st.subheader("Actual vs Predicted Values")
            
            # Create a melted dataframe for Plotly
            df_melted = pd.melt(df, id_vars=['Bag Plus Plant'], value_vars=['Mar-Actual', 'Mar Pred1', 'Mar Pred2'],
                                var_name='Measurement Type', value_name='Value')
            
            fig = px.bar(df_melted, x='Bag Plus Plant', y='Value', color='Measurement Type', barmode='group',
                        color_discrete_map={'Mar-Actual': '#2ecc71', 'Mar Pred1': '#3498db', 'Mar Pred2': '#9b59b6'},
                        title='Comparison of Actual vs Predicted Consumption by Cement Bag Type')
            
            fig.update_layout(xaxis_title='Cement Bag Type', yaxis_title='Consumption',
                            legend_title='Data Type', template='plotly_white')
            
            # Rotate x-axis labels if there are many cement bag types
            if len(df) > 5:
                fig.update_layout(xaxis_tickangle=-45)
                
            st.plotly_chart(fig, use_container_width=True)
            
            # 2. Error Analysis
            st.subheader("Error Analysis")
            
            # Create subplots for error distribution
            error_fig = make_subplots(rows=1, cols=2, subplot_titles=("Model 1 Error Distribution", "Model 2 Error Distribution"))
            
            # Model 1 error histogram
            error_fig.add_trace(
                go.Histogram(x=df['Error_Model1'], name='Model 1 Error', marker_color='#3498db'),
                row=1, col=1
            )
            
            # Model 2 error histogram
            error_fig.add_trace(
                go.Histogram(x=df['Error_Model2'], name='Model 2 Error', marker_color='#9b59b6'),
                row=1, col=2
            )
            
            error_fig.update_layout(height=500, title_text="Error Distribution Comparison",
                                    template='plotly_white')
            
            st.plotly_chart(error_fig, use_container_width=True)
            
            # 3. Scatter plots of Actual vs Predicted
            scatter_fig = make_subplots(rows=1, cols=2, subplot_titles=("Model 1: Actual vs Predicted", "Model 2: Actual vs Predicted"),
                                        specs=[[{"type": "scatter"}, {"type": "scatter"}]])
            
            # Add identity line (perfect prediction)
            max_val = max(df['Mar-Actual'].max(), df['Mar Pred1'].max(), df['Mar Pred2'].max())
            min_val = min(df['Mar-Actual'].min(), df['Mar Pred1'].min(), df['Mar Pred2'].min())
            
            # Add identity line to both subplots
            for col in [1, 2]:
                scatter_fig.add_trace(
                    go.Scatter(x=[min_val, max_val], y=[min_val, max_val], mode='lines', name='Perfect Prediction',
                              line=dict(color='rgba(0,0,0,0.5)', dash='dash'), showlegend=col==1),
                    row=1, col=col
                )
            
            # Model 1 scatter plot
            scatter_fig.add_trace(
                go.Scatter(x=df['Mar-Actual'], y=df['Mar Pred1'], mode='markers', name='Model 1',
                          marker=dict(color='#3498db', size=10)),
                row=1, col=1
            )
            
            # Model 2 scatter plot
            scatter_fig.add_trace(
                go.Scatter(x=df['Mar-Actual'], y=df['Mar Pred2'], mode='markers', name='Model 2',
                          marker=dict(color='#9b59b6', size=10)),
                row=1, col=2
            )
            
            scatter_fig.update_layout(height=500, title_text="Actual vs Predicted Scatter Plots",
                                    xaxis_title="Actual Values", yaxis_title="Predicted Values",
                                    xaxis2_title="Actual Values", yaxis2_title="Predicted Values",
                                    template='plotly_white')
            
            st.plotly_chart(scatter_fig, use_container_width=True)
            
            # 4. Percentage Error by Cement Bag Type
            st.subheader("Percentage Error by Cement Bag Type")
            
            percent_error_df = pd.DataFrame({
                'Bag Plus Plant': df['Bag Plus Plant'],
                'Model 1 Error (%)': df['Error_Percent_Model1'],
                'Model 2 Error (%)': df['Error_Percent_Model2']
            })
            
            percent_error_melted = pd.melt(percent_error_df, id_vars=['Bag Plus Plant'], 
                                        value_vars=['Model 1 Error (%)', 'Model 2 Error (%)'],
                                        var_name='Model', value_name='Percentage Error')
            
            percent_fig = px.bar(percent_error_melted, x='Bag Plus Plant', y='Percentage Error', color='Model',
                                barmode='group', title='Percentage Error Comparison by Cement Bag Type',
                                color_discrete_map={'Model 1 Error (%)': '#3498db', 'Model 2 Error (%)': '#9b59b6'})
            
            percent_fig.update_layout(xaxis_title='Cement Bag Type', yaxis_title='Percentage Error (%)',
                                    legend_title='Model', template='plotly_white')
            
            if len(df) > 5:
                percent_fig.update_layout(xaxis_tickangle=-45)
                
            st.plotly_chart(percent_fig, use_container_width=True)
            
            # 5. Cumulative Error Analysis
            st.subheader("Cumulative Error Analysis")
            
            # Sort dataframes by actual consumption
            df_sorted = df.sort_values('Mar-Actual')
            
            # Calculate cumulative sums
            df_sorted['Cumulative_Actual'] = df_sorted['Mar-Actual'].cumsum()
            df_sorted['Cumulative_Pred1'] = df_sorted['Mar Pred1'].cumsum()
            df_sorted['Cumulative_Pred2'] = df_sorted['Mar Pred2'].cumsum()
            
            # Calculate cumulative errors
            df_sorted['Cumulative_Error_Model1'] = df_sorted['Cumulative_Actual'] - df_sorted['Cumulative_Pred1']
            df_sorted['Cumulative_Error_Model2'] = df_sorted['Cumulative_Actual'] - df_sorted['Cumulative_Pred2']
            
            # Create the figure
            cum_fig = go.Figure()
            
            cum_fig.add_trace(go.Scatter(x=df_sorted['Bag Plus Plant'], y=df_sorted['Cumulative_Error_Model1'], 
                                        mode='lines+markers', name='Model 1 Cumulative Error',
                                        line=dict(color='#3498db', width=2)))
            
            cum_fig.add_trace(go.Scatter(x=df_sorted['Bag Plus Plant'], y=df_sorted['Cumulative_Error_Model2'], 
                                        mode='lines+markers', name='Model 2 Cumulative Error',
                                        line=dict(color='#9b59b6', width=2)))
            
            # Add horizontal line at y=0
            cum_fig.add_hline(y=0, line_width=1, line_dash="dash", line_color="black")
            
            cum_fig.update_layout(title='Cumulative Error Analysis', xaxis_title='Cement Bag Type (Sorted by Actual Consumption)',
                                yaxis_title='Cumulative Error', template='plotly_white')
            
            if len(df) > 5:
                cum_fig.update_layout(xaxis_tickangle=-45)
                
            st.plotly_chart(cum_fig, use_container_width=True)
            
            # 6. Radar Chart for Model Comparison
            st.subheader("Radar Chart: Model Performance Comparison")
            
            # Select metrics for radar chart (normalize them for better visualization)
            radar_metrics = ['MAE', 'RMSE', 'MAPE', 'R²', 'Within 5% Error (%)', 'Within 10% Error (%)']
            
            # Create a dataframe for radar chart
            radar_df = pd.DataFrame({
                'Metric': radar_metrics,
                'Model 1': [metrics_model1[m] for m in radar_metrics],
                'Model 2': [metrics_model2[m] for m in radar_metrics]
            })
            
            # Normalize metrics (invert for metrics where lower is better)
            for metric in radar_metrics:
                if metric in ['R²', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                    # Higher is better - normalize to 0-1 range
                    max_val = max(radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],
                                radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0])
                    
                    if max_val != 0:
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val
                else:
                    # Lower is better - invert and normalize to 0-1 range
                    max_val = max(radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],
                                radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0])
                    
                    if max_val != 0:
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val)
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val)
            
            # Create radar chart
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
                    )),
                showlegend=True,
                title='Model Performance Radar Chart (Higher is Better for All Metrics)',
                template='plotly_white'
            )
            
            st.plotly_chart(radar_fig, use_container_width=True)
            
            # 7. Download analysis as Excel
            st.subheader("Download Analysis Results")
            
            # Create a BytesIO object
            output = io.BytesIO()
            
            # Create an Excel writer using the BytesIO object
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                # Write original data
                df.to_excel(writer, sheet_name='Original Data', index=False)
                
                # Write metrics comparison
                comparison_df.to_excel(writer, sheet_name='Metrics Comparison', index=False)
                
                # Format the excel file
                workbook = writer.book
                worksheet = writer.sheets['Metrics Comparison']
                
                # Add formats
                better_format = workbook.add_format({'bg_color': '#d4f1dd'})
                worse_format = workbook.add_format({'bg_color': '#f8d7da'})
                
                # Apply formats based on better model
                for row_num, model in enumerate(better_model, start=1):
                    if model == 'Model 1':
                        worksheet.conditional_format(row_num, 1, row_num, 1, {'type': 'no_blanks', 'format': better_format})
                        worksheet.conditional_format(row_num, 2, row_num, 2, {'type': 'no_blanks', 'format': worse_format})
                    elif model == 'Model 2':
                        worksheet.conditional_format(row_num, 1, row_num, 1, {'type': 'no_blanks', 'format': worse_format})
                        worksheet.conditional_format(row_num, 2, row_num, 2, {'type': 'no_blanks', 'format': better_format})
            
            # Get the value of the BytesIO buffer
            excel_data = output.getvalue()
            
            # Provide download button
            st.download_button(
                label="Download Analysis as Excel",
                data=excel_data,
                file_name="cement_model_comparison_analysis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
        else:
            st.error(f"Required columns not found. Please ensure your Excel file has these columns: {', '.join(required_columns)}")
    
    except Exception as e:
        st.error(f"Error processing the file: {str(e)}")
else:
    # Display a sample template when no file is uploaded
    st.info("Please upload an Excel file with the following columns: 'Bag Plus Plant', 'Mar-Actual', 'Mar Pred1', 'Mar Pred2'")
    
    # Show sample data structure
    sample_df = pd.DataFrame({
        'Bag Plus Plant': ['Cement Type A - Plant 1', 'Cement Type B - Plant 2', 'Cement Type C - Plant 1'],
        'Mar-Actual': [1500, 2000, 1200],
        'Mar Pred1': [1450, 2100, 1250],
        'Mar Pred2': [1530, 1950, 1180]
    })
    
    st.write("Sample data structure:")
    st.dataframe(sample_df)

# Add footer
st.markdown("""
<div style="text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; border-radius: 5px;">
    <p style="color: #7f8c8d;">Cement Consumption Model Comparison Dashboard</p>
</div>
""", unsafe_allow_html=True)
