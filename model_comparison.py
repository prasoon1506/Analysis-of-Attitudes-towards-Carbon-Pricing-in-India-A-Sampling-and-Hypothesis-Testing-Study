import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
from scipy.stats import pearsonr, spearmanr
from sklearn.metrics import mean_absolute_error, mean_squared_error, r2_score
import io

# Set page configuration
st.set_page_config(layout="wide", page_title="Cement Consumption Model Comparison")

# Custom CSS
st.markdown("""<style>
.main {background-color: #f8f9fa;}
h1, h2, h3 {color: #2c3e50;}
.stButton>button {background-color: #3498db; color: white;}
.metric-card {background-color: white; border-radius: 5px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1); padding: 20px; margin: 10px 0;}
.highlight {font-weight: bold; color: #27ae60;}
.month-section {border: 1px solid #e0e0e0; border-radius: 10px; padding: 15px; margin: 10px 0; background-color: #f9f9f9;}
</style>""", unsafe_allow_html=True)

# Page title
st.title("Cement Bag Consumption Model Comparison Dashboard")

# Dashboard description
st.markdown("""<div style="background-color: #f0f5ff; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db;">
<h3 style="margin-top: 0;">Model Comparison</h3>
<p><strong>Model 1:</strong> Neural Network Algorithm</p>
<p><strong>Model 2:</strong> Ensemble Algorithm (Holt-Winters + Trend-Based + Random-Forest)</p>
<p>This dashboard provides a comprehensive comparison between these two prediction models for cement bag consumption against actual values for two months.</p>
</div>""", unsafe_allow_html=True)

# Define metric calculation function
def calculate_metrics(actual, predicted):
    mae = mean_absolute_error(actual, predicted)
    mse = mean_squared_error(actual, predicted)
    rmse = np.sqrt(mse)
    mape = np.mean(np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual))) * 100
    wmape = np.sum(np.abs(actual - predicted)) / np.sum(np.maximum(np.ones(len(actual)), actual)) * 100
    r2 = r2_score(actual, predicted)
    pearson_corr, _ = pearsonr(actual, predicted)
    spearman_corr, _ = spearmanr(actual, predicted)
    
    percent_errors = np.abs((actual - predicted) / np.maximum(np.ones(len(actual)), actual)) * 100
    under_predictions = np.sum(predicted < actual)
    over_predictions = np.sum(predicted > actual)
    perfect_predictions = np.sum(predicted == actual)
    
    within_1_percent = np.sum(percent_errors <= 1)
    within_1_percent_ratio = within_1_percent / len(actual) * 100
    within_3_percent = np.sum(percent_errors <= 3)
    within_3_percent_ratio = within_3_percent / len(actual) * 100
    within_5_percent = np.sum(percent_errors <= 5)
    within_5_percent_ratio = within_5_percent / len(actual) * 100
    within_10_percent = np.sum(percent_errors <= 10)
    within_10_percent_ratio = within_10_percent / len(actual) * 100
    
    total_actual = np.sum(actual)
    total_predicted = np.sum(predicted)
    abs_total_deviation = abs(total_actual - total_predicted)
    total_deviation_percent = (abs_total_deviation / total_actual) * 100 if total_actual > 0 else 0
    
    bias = np.mean(predicted - actual)
    bias_percent = (bias / np.mean(actual)) * 100 if np.mean(actual) > 0 else 0
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

# Function to process a month's data and create visualizations
def process_month_data(df, month_prefix, container):
    with container:
        # Calculate errors
        df[f'Error_Model1'] = df[f'{month_prefix}-Actual'] - df[f'{month_prefix} Pred1']
        df[f'Error_Model2'] = df[f'{month_prefix}-Actual'] - df[f'{month_prefix} Pred2']
        df[f'Error_Percent_Model1'] = np.abs(df[f'Error_Model1'] / np.maximum(np.ones(len(df)), df[f'{month_prefix}-Actual'])) * 100
        df[f'Error_Percent_Model2'] = np.abs(df[f'Error_Model2'] / np.maximum(np.ones(len(df)), df[f'{month_prefix}-Actual'])) * 100
        
        # Calculate metrics
        metrics_model1 = calculate_metrics(df[f'{month_prefix}-Actual'], df[f'{month_prefix} Pred1'])
        metrics_model2 = calculate_metrics(df[f'{month_prefix}-Actual'], df[f'{month_prefix} Pred2'])
        
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
        
        # Model comparison summary
        st.subheader("Model Comparison Summary")
        comparison_df = pd.DataFrame({
            'Metric': [k for k in metrics_model1.keys() if k != 'Percent Errors'],
            'Model 1': [metrics_model1[k] for k in metrics_model1.keys() if k != 'Percent Errors'],
            'Model 2': [metrics_model2[k] for k in metrics_model1.keys() if k != 'Percent Errors']
        })
        
        better_model = []
        for metric in comparison_df['Metric']:
            if metric in ['R²', 'Pearson Correlation', 'Spearman Correlation', 'Perfect Predictions', 'Within 1% Error (%)', 
                         'Within 3% Error (%)', 'Within 5% Error (%)', 'Within 10% Error (%)']:
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
        st.dataframe(comparison_df.style.apply(lambda x: ['background-color: #d4f1dd' if v == "Model 1" 
                                                        else 'background-color: #d1e7f0' if v == "Model 2" 
                                                        else '' for v in x], subset=['Better Model']))
        
        # Calculate winner
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
        
        # 1. Actual vs Predicted Values bar chart
        st.subheader("Actual vs Predicted Values")
        df_melted = pd.melt(
            df, 
            id_vars=['Bag Plus Plant'], 
            value_vars=[f'{month_prefix}-Actual', f'{month_prefix} Pred1', f'{month_prefix} Pred2'],
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
                f'{month_prefix}-Actual': '#2ecc71', 
                f'{month_prefix} Pred1': '#3498db', 
                f'{month_prefix} Pred2': '#9b59b6'
            },
            title=f'Comparison of Actual vs Predicted Consumption by Cement Bag Type ({month_prefix})'
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
        
        # 2. Error Analysis histogram
        st.subheader("Error Analysis")
        error_fig = make_subplots(
            rows=1, 
            cols=2, 
            subplot_titles=(f"Model 1 Error Distribution ({month_prefix})", f"Model 2 Error Distribution ({month_prefix})")
        )
        error_fig.add_trace(
            go.Histogram(x=df[f'Error_Model1'], name='Model 1 Error', marker_color='#3498db'),
            row=1, col=1
        )
        error_fig.add_trace(
            go.Histogram(x=df[f'Error_Model2'], name='Model 2 Error', marker_color='#9b59b6'),
            row=1, col=2
        )
        error_fig.update_layout(
            height=500, 
            title_text=f"Error Distribution Comparison ({month_prefix})",
            template='plotly_white'
        )
        st.plotly_chart(error_fig, use_container_width=True)
        
        # 3. Scatter plots
        scatter_fig = make_subplots(
            rows=1, 
            cols=2, 
            subplot_titles=(f"Model 1: Actual vs Predicted ({month_prefix})", f"Model 2: Actual vs Predicted ({month_prefix})"),
            specs=[[{"type": "scatter"}, {"type": "scatter"}]]
        )
        max_val = max(df[f'{month_prefix}-Actual'].max(), df[f'{month_prefix} Pred1'].max(), df[f'{month_prefix} Pred2'].max())
        min_val = min(df[f'{month_prefix}-Actual'].min(), df[f'{month_prefix} Pred1'].min(), df[f'{month_prefix} Pred2'].min())
        
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
                x=df[f'{month_prefix}-Actual'], 
                y=df[f'{month_prefix} Pred1'], 
                mode='markers', 
                name='Model 1',
                marker=dict(color='#3498db', size=10)
            ),
            row=1, col=1
        )
        
        scatter_fig.add_trace(
            go.Scatter(
                x=df[f'{month_prefix}-Actual'], 
                y=df[f'{month_prefix} Pred2'], 
                mode='markers', 
                name='Model 2',
                marker=dict(color='#9b59b6', size=10)
            ),
            row=1, col=2
        )
        
        scatter_fig.update_layout(
            height=500, 
            title_text=f"Actual vs Predicted Scatter Plots ({month_prefix})",
            xaxis_title="Actual Values", 
            yaxis_title="Predicted Values",
            xaxis2_title="Actual Values", 
            yaxis2_title="Predicted Values",
            template='plotly_white'
        )
        st.plotly_chart(scatter_fig, use_container_width=True)
        
        # 4. Percentage Error by Cement Bag Type
        st.subheader("Percentage Error by Cement Bag Type")
        percent_error_df = pd.DataFrame({
            'Bag Plus Plant': df['Bag Plus Plant'],
            'Model 1 Error (%)': df[f'Error_Percent_Model1'],
            'Model 2 Error (%)': df[f'Error_Percent_Model2']
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
            title=f'Percentage Error Comparison by Cement Bag Type ({month_prefix})',
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
        
        # 5. Error heat map
        st.subheader("Prediction Accuracy Heat Map")
        heatmap_data = pd.DataFrame({
            'Bag Plus Plant': df['Bag Plus Plant'],
            'Neural Network Error (%)': df[f'Error_Percent_Model1'],
            'Ensemble Algorithm Error (%)': df[f'Error_Percent_Model2']
        })
        
        heatmap_pivot = heatmap_data.set_index('Bag Plus Plant')
        heatmap_fig = px.imshow(
            heatmap_pivot.T,
            text_auto='.1f',
            aspect="auto",
            color_continuous_scale='RdYlGn_r',
            title=f'Prediction Error Heat Map (%) - {month_prefix}',
            labels=dict(x="Cement Bag Type", y="Model", color="Error (%)")
        )
        
        heatmap_fig.update_layout(height=400, template='plotly_white')
        
        if len(df) > 5:
            heatmap_fig.update_layout(xaxis_tickangle=-45)
        
        st.plotly_chart(heatmap_fig, use_container_width=True)
        
        # 6. Radar chart
        st.subheader("Radar Chart: Model Performance Comparison")
        radar_metrics = ['MAE', 'RMSE', 'MAPE', 'R²', 'Within 5% Error (%)', 'Within 10% Error (%)']
        radar_df = pd.DataFrame({
            'Metric': radar_metrics,
            'Model 1': [metrics_model1[m] for m in radar_metrics],
            'Model 2': [metrics_model2[m] for m in radar_metrics]
        })
        
        for metric in radar_metrics:
            if metric in ['R²', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                # Higher is better, normalize to 0-1 range
                max_val = max(
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0]
                )
                if max_val != 0:
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val
                    radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val
            else:
                # Lower is better, invert so higher is better for radar chart display
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
            polar=dict(radialaxis=dict(visible=True, range=[0, 1])),
            showlegend=True,
            title=f'Model Performance Radar Chart ({month_prefix}) - Higher is Better for All Metrics',
            template='plotly_white'
        )
        
        st.plotly_chart(radar_fig, use_container_width=True)
        
        # Return the processed dataframe and metrics for possible combined analysis
        return df, metrics_model1, metrics_model2

# Create file upload fields
st.header("Upload Excel Files")
col1, col2 = st.columns(2)

with col1:
    st.subheader("March File")
    march_file = st.file_uploader("Upload March Excel File", type=["xlsx", "xls"], key="march")

with col2:
    st.subheader("April File")
    april_file = st.file_uploader("Upload April Excel File", type=["xlsx", "xls"], key="april")

# Process the files when both are uploaded
if march_file and april_file:
    try:
        # Read files
        df_march = pd.read_excel(march_file)
        df_april = pd.read_excel(april_file)
        
        # Validate March file
        march_required_columns = ["Bag Plus Plant", "Mar-Actual", "Mar Pred1", "Mar Pred2"]
        march_valid = all(col in df_march.columns for col in march_required_columns)
        
        # Validate April file
        april_required_columns = ["Bag Plus Plant", "Apr-Actual", "Apr Pred1", "Apr Pred2"]
        april_valid = all(col in df_april.columns for col in april_required_columns)
        
        if march_valid and april_valid:
            # Display raw data preview
            with st.expander("Raw Data Preview"):
                st.subheader("March Data")
                st.dataframe(df_march)
                st.subheader("April Data")
                st.dataframe(df_april)
            
            # Process each month in separate containers
            st.markdown("<hr>", unsafe_allow_html=True)
            
            # March section
            st.markdown('<div class="month-section">', unsafe_allow_html=True)
            st.header("March Analysis")
            march_container = st.container()
            march_df, march_metrics1, march_metrics2 = process_month_data(df_march, "Mar", march_container)
            st.markdown('</div>', unsafe_allow_html=True)
            
            # April section
            st.markdown('<div class="month-section">', unsafe_allow_html=True)
            st.header("April Analysis")
            april_container = st.container()
            april_df, april_metrics1, april_metrics2 = process_month_data(df_april, "Apr", april_container)
            st.markdown('</div>', unsafe_allow_html=True)
            
            # Month-to-Month Comparison
            st.markdown("<hr>", unsafe_allow_html=True)
            st.header("Month-to-Month Comparison")
            
            # Metric improvement comparison
            improvement_metrics = ['MAE', 'RMSE', 'MAPE', 'WMAPE', 'R²', 'Within 5% Error (%)', 'Within 10% Error (%)']
            
            improvement_df = pd.DataFrame({
                'Metric': improvement_metrics,
                'Model 1 March': [march_metrics1[m] for m in improvement_metrics],
                'Model 1 April': [april_metrics1[m] for m in improvement_metrics],
                'Model 1 Change (%)': [(april_metrics1[m] - march_metrics1[m]) / march_metrics1[m] * 100 if march_metrics1[m] != 0 else 0 for m in improvement_metrics],
                'Model 2 March': [march_metrics2[m] for m in improvement_metrics],
                'Model 2 April': [april_metrics2[m] for m in improvement_metrics],
                'Model 2 Change (%)': [(april_metrics2[m] - march_metrics2[m]) / march_metrics2[m] * 100 if march_metrics2[m] != 0 else 0 for m in improvement_metrics]
            })
            
            # Format the improvement dataframe
            for i, metric in enumerate(improvement_metrics):
                if metric in ['R²', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                    # Higher is better
                    improvement_df.loc[i, 'Model 1 Improved'] = "Yes" if improvement_df.loc[i, 'Model 1 Change (%)'] > 0 else "No"
                    improvement_df.loc[i, 'Model 2 Improved'] = "Yes" if improvement_df.loc[i, 'Model 2 Change (%)'] > 0 else "No"
                else:
                    # Lower is better
                    improvement_df.loc[i, 'Model 1 Improved'] = "Yes" if improvement_df.loc[i, 'Model 1 Change (%)'] < 0 else "No"
                    improvement_df.loc[i, 'Model 2 Improved'] = "Yes" if improvement_df.loc[i, 'Model 2 Change (%)'] < 0 else "No"
            
            st.subheader("Metric Changes from March to April")
            st.dataframe(improvement_df[['Metric', 'Model 1 March', 'Model 1 April', 'Model 1 Change (%)', 'Model 1 Improved', 
                                        'Model 2 March', 'Model 2 April', 'Model 2 Change (%)', 'Model 2 Improved']])
            
            # Month-to-month performance comparison visualization
            st.subheader("Month-to-Month Model Performance")
            
            # Prepare data for the comparison chart
            m2m_data = pd.DataFrame({
                'Month': ['March', 'March', 'April', 'April'],
                'Model': ['Model 1', 'Model 2', 'Model 1', 'Model 2'],
                'MAPE': [march_metrics1['MAPE'], march_metrics2['MAPE'], april_metrics1['MAPE'], april_metrics2['MAPE']],
                'RMSE': [march_metrics1['RMSE'], march_metrics2['RMSE'], april_metrics1['RMSE'], april_metrics2['RMSE']],
                'R²': [march_metrics1['R²'], march_metrics2['R²'], april_metrics1['R²'], april_metrics2['R²']],
                'Within 5% Error (%)': [march_metrics1['Within 5% Error (%)'], march_metrics2['Within 5% Error (%)'], 
                                      april_metrics1['Within 5% Error (%)'], april_metrics2['Within 5% Error (%)']],
            })
            
            # Create comparison charts
            tab1, tab2, tab3, tab4 = st.tabs(["MAPE", "RMSE", "R²", "Within 5% Error"])
            
            with tab1:
                fig = px.bar(
                    m2m_data, 
                    x='Month', 
                    y='MAPE', 
                    color='Model', 
                    barmode='group',
                    color_discrete_map={'Model 1': '#3498db', 'Model 2': '#9b59b6'},
                    title='MAPE Comparison (Lower is Better)'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with tab2:
                fig = px.bar(
                    m2m_data, 
                    x='Month', 
                    y='RMSE', 
                    color='Model', 
                    barmode='group',
                    color_discrete_map={'Model 1': '#3498db', 'Model 2': '#9b59b6'},
                    title='RMSE Comparison (Lower is Better)'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with tab3:
                fig = px.bar(
                    m2m_data, 
                    x='Month', 
                    y='R²', 
                    color='Model', 
                    barmode='group',
                    color_discrete_map={'Model 1': '#3498db', 'Model 2': '#9b59b6'},
                    title='R² Comparison (Higher is Better)'
                )
                st.plotly_chart(fig, use_container_width=True)
            
            with tab4:
                fig = px.bar(
                    m2m_data, 
                    x='Month', 
                    y='Within 5% Error (%)', 
                    color='Model', 
                    barmode='group',
                    color_discrete_map={'Model 1': '#3498db', 'Model 2': '#9b59b6'},
                    title='Within 5% Error Comparison (Higher is Better)'
                )
                st.plotly_chart(fig, use_container_width=True)
            

else:
    st.info("Please upload an Excel file with the following columns: 'Bag Plus Plant', 'Mar-Actual', 'Mar Pred1', 'Mar Pred2'")
    sample_df = pd.DataFrame({'Bag Plus Plant': ['Cement Type A - Plant 1', 'Cement Type B - Plant 2', 'Cement Type C - Plant 1'],'Mar-Actual': [1500, 2000, 1200],'Mar Pred1': [1450, 2100, 1250],'Mar Pred2': [1530, 1950, 1180]})
    st.write("Sample data structure:")
    st.dataframe(sample_df)
st.markdown("""<div style="text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; border-radius: 5px;"><p style="color: #7f8c8d;">Cement Consumption Model Comparison Dashboard</p></div>""", unsafe_allow_html=True)
