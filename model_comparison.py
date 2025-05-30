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
st.markdown("""<style>.main {background-color: #f8f9fa;}h1, h2, h3 {color: #2c3e50;}.stButton>button {background-color: #3498db;color: white;}.metric-card {background-color: white;border-radius: 5px;box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);padding: 20px;margin: 10px 0;}.highlight {font-weight: bold;color: #27ae60;}</style>""", unsafe_allow_html=True)
st.title("Cement Bag Consumption Model Comparison Dashboard")
st.markdown("""<div style="background-color: #f0f5ff; padding: 15px; border-radius: 10px; border-left: 5px solid #3498db;"><h3 style="margin-top: 0;">Model Comparison</h3><p><strong>Model 1:</strong> Neural Network Algorithm</p><p><strong>Model 2:</strong> Ensemble Algorithm (Holt-Winters + Trend-Based + Random-Forest)</p><p>This dashboard provides a comprehensive comparison between these two prediction models for cement bag consumption against actual values for March.</p></div>""", unsafe_allow_html=True)
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
    return {'MAE': mae,'MSE': mse,'RMSE': rmse,'MAPE': mape,'WMAPE': wmape,'R²': r2,'Pearson Correlation': pearson_corr,'Spearman Correlation': spearman_corr,'Under Predictions': under_predictions,'Over Predictions': over_predictions,'Perfect Predictions': perfect_predictions,'Within 1% Error (%)': within_1_percent_ratio,'Within 3% Error (%)': within_3_percent_ratio,'Within 5% Error (%)': within_5_percent_ratio,'Within 10% Error (%)': within_10_percent_ratio,'Total Deviation (%)': total_deviation_percent,'Bias': bias,'Bias (%)': bias_percent,'Tracking Signal': tracking_signal,'Percent Errors': percent_errors}
st.subheader("Upload Excel File")
uploaded_file = st.file_uploader("Upload your Excel file with cement bag data", type=["xlsx", "xls"])
if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        with st.expander("Raw Data Preview"):
            st.dataframe(df)
        required_columns = ["Bag Plus Plant", "Mar-Actual", "Mar Pred1", "Mar Pred2"]
        if all(col in df.columns for col in required_columns):
            df['Error_Model1'] = df['Mar-Actual'] - df['Mar Pred1']
            df['Error_Model2'] = df['Mar-Actual'] - df['Mar Pred2']
            df['Error_Percent_Model1'] = np.abs(df['Error_Model1'] / np.maximum(np.ones(len(df)), df['Mar-Actual'])) * 100
            df['Error_Percent_Model2'] = np.abs(df['Error_Model2'] / np.maximum(np.ones(len(df)), df['Mar-Actual'])) * 100
            metrics_model1 = calculate_metrics(df['Mar-Actual'], df['Mar Pred1'])
            metrics_model2 = calculate_metrics(df['Mar-Actual'], df['Mar Pred2'])
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
            st.subheader("Model Comparison Summary")
            comparison_df = pd.DataFrame({'Metric': [k for k in metrics_model1.keys() if k != 'Percent Errors'],'Model 1': [metrics_model1[k] for k in metrics_model1.keys() if k != 'Percent Errors'],'Model 2': [metrics_model2[k] for k in metrics_model1.keys() if k != 'Percent Errors']})
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
                    if metrics_model1[metric] < metrics_model2[metric]:
                        better_model.append("Model 1")
                    elif metrics_model2[metric] < metrics_model1[metric]:
                        better_model.append("Model 2")
                    else:
                        better_model.append("Equal")
            comparison_df['Better Model'] = better_model
            st.dataframe(comparison_df.style.apply(lambda x: ['background-color: #d4f1dd' if v == "Model 1" else 'background-color: #d1e7f0' if v == "Model 2" else '' for v in x], subset=['Better Model']))
            model1_wins = sum(1 for model in better_model if model == "Model 1")
            model2_wins = sum(1 for model in better_model if model == "Model 2")
            equal_metrics = sum(1 for model in better_model if model == "Equal")
            winner = "Model 1" if model1_wins > model2_wins else "Model 2" if model2_wins > model1_wins else "Both models perform equally"
            st.markdown(f"""<div class="metric-card"><h3>Overall Winner: <span class="highlight">{winner}</span></h3><p>Model 1 better in {model1_wins} metrics</p><p>Model 2 better in {model2_wins} metrics</p><p>Equal performance in {equal_metrics} metrics</p></div>""", unsafe_allow_html=True)
            st.header("Visualizations")
            st.subheader("Actual vs Predicted Values")
            df_melted = pd.melt(df, id_vars=['Bag Plus Plant'], value_vars=['Mar-Actual', 'Mar Pred1', 'Mar Pred2'],var_name='Measurement Type', value_name='Value')
            fig = px.bar(df_melted, x='Bag Plus Plant', y='Value', color='Measurement Type', barmode='group',color_discrete_map={'Mar-Actual': '#2ecc71', 'Mar Pred1': '#3498db', 'Mar Pred2': '#9b59b6'},title='Comparison of Actual vs Predicted Consumption by Cement Bag Type')
            fig.update_layout(xaxis_title='Cement Bag Type', yaxis_title='Consumption',legend_title='Data Type', template='plotly_white')
            if len(df) > 5:
                fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(fig, use_container_width=True)
            st.subheader("Error Analysis")
            error_fig = make_subplots(rows=1, cols=2, subplot_titles=("Model 1 Error Distribution", "Model 2 Error Distribution"))
            error_fig.add_trace(go.Histogram(x=df['Error_Model1'], name='Model 1 Error', marker_color='#3498db'),row=1, col=1)
            error_fig.add_trace(go.Histogram(x=df['Error_Model2'], name='Model 2 Error', marker_color='#9b59b6'),row=1, col=2)
            error_fig.update_layout(height=500, title_text="Error Distribution Comparison",template='plotly_white')
            st.plotly_chart(error_fig, use_container_width=True)
            scatter_fig = make_subplots(rows=1, cols=2, subplot_titles=("Model 1: Actual vs Predicted", "Model 2: Actual vs Predicted"),specs=[[{"type": "scatter"}, {"type": "scatter"}]])
            max_val = max(df['Mar-Actual'].max(), df['Mar Pred1'].max(), df['Mar Pred2'].max())
            min_val = min(df['Mar-Actual'].min(), df['Mar Pred1'].min(), df['Mar Pred2'].min())
            for col in [1, 2]:
                scatter_fig.add_trace(go.Scatter(x=[min_val, max_val], y=[min_val, max_val], mode='lines', name='Perfect Prediction',line=dict(color='rgba(0,0,0,0.5)', dash='dash'), showlegend=col==1),row=1, col=col)
            scatter_fig.add_trace(go.Scatter(x=df['Mar-Actual'], y=df['Mar Pred1'], mode='markers', name='Model 1',marker=dict(color='#3498db', size=10)),row=1, col=1)
            scatter_fig.add_trace(go.Scatter(x=df['Mar-Actual'], y=df['Mar Pred2'], mode='markers', name='Model 2',marker=dict(color='#9b59b6', size=10)),row=1, col=2)
            scatter_fig.update_layout(height=500, title_text="Actual vs Predicted Scatter Plots",xaxis_title="Actual Values", yaxis_title="Predicted Values",xaxis2_title="Actual Values", yaxis2_title="Predicted Values",template='plotly_white')
            st.plotly_chart(scatter_fig, use_container_width=True)
            st.subheader("Percentage Error by Cement Bag Type")
            percent_error_df = pd.DataFrame({'Bag Plus Plant': df['Bag Plus Plant'],'Model 1 Error (%)': df['Error_Percent_Model1'],'Model 2 Error (%)': df['Error_Percent_Model2']})
            percent_error_melted = pd.melt(percent_error_df, id_vars=['Bag Plus Plant'],value_vars=['Model 1 Error (%)', 'Model 2 Error (%)'],var_name='Model', value_name='Percentage Error')
            percent_fig = px.bar(percent_error_melted, x='Bag Plus Plant', y='Percentage Error', color='Model',barmode='group', title='Percentage Error Comparison by Cement Bag Type',color_discrete_map={'Model 1 Error (%)': '#3498db', 'Model 2 Error (%)': '#9b59b6'}) 
            percent_fig.update_layout(xaxis_title='Cement Bag Type', yaxis_title='Percentage Error (%)',legend_title='Model', template='plotly_white')
            if len(df) > 5:
                percent_fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(percent_fig, use_container_width=True)
            st.subheader("Cumulative Error Analysis")
            df_sorted = df.sort_values('Mar-Actual')
            df_sorted['Cumulative_Actual'] = df_sorted['Mar-Actual'].cumsum()
            df_sorted['Cumulative_Pred1'] = df_sorted['Mar Pred1'].cumsum()
            df_sorted['Cumulative_Pred2'] = df_sorted['Mar Pred2'].cumsum()
            df_sorted['Cumulative_Error_Model1'] = df_sorted['Cumulative_Actual'] - df_sorted['Cumulative_Pred1']
            df_sorted['Cumulative_Error_Model2'] = df_sorted['Cumulative_Actual'] - df_sorted['Cumulative_Pred2']
            cum_fig = go.Figure()
            cum_fig.add_trace(go.Scatter(x=df_sorted['Bag Plus Plant'], y=df_sorted['Cumulative_Error_Model1'],mode='lines+markers', name='Model 1 Cumulative Error',line=dict(color='#3498db', width=2)))  
            cum_fig.add_trace(go.Scatter(x=df_sorted['Bag Plus Plant'], y=df_sorted['Cumulative_Error_Model2'], mode='lines+markers', name='Model 2 Cumulative Error',line=dict(color='#9b59b6', width=2)))
            cum_fig.add_hline(y=0, line_width=1, line_dash="dash", line_color="black")
            cum_fig.update_layout(title='Cumulative Error Analysis', xaxis_title='Cement Bag Type (Sorted by Actual Consumption)',yaxis_title='Cumulative Error', template='plotly_white')
            if len(df) > 5:
                cum_fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(cum_fig, use_container_width=True)
            st.subheader("Radar Chart: Model Performance Comparison")
            radar_metrics = ['MAE', 'RMSE', 'MAPE', 'R²', 'Within 5% Error (%)', 'Within 10% Error (%)']
            radar_df = pd.DataFrame({'Metric': radar_metrics,'Model 1': [metrics_model1[m] for m in radar_metrics],'Model 2': [metrics_model2[m] for m in radar_metrics]})
            for metric in radar_metrics:
                if metric in ['R²', 'Within 5% Error (%)', 'Within 10% Error (%)']:
                    max_val = max(radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0])
                    if max_val != 0:
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val
                else:
                    max_val = max(radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0],radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0])
                    if max_val != 0:
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 1'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 1'].iloc[0] / max_val)
                        radar_df.loc[radar_df['Metric'] == metric, 'Model 2'] = 1 - (radar_df.loc[radar_df['Metric'] == metric, 'Model 2'].iloc[0] / max_val)
            radar_fig = go.Figure()
            radar_fig.add_trace(go.Scatterpolar(r=radar_df['Model 1'].values,theta=radar_df['Metric'].values,fill='toself',name='Model 1',line_color='#3498db'))
            radar_fig.add_trace(go.Scatterpolar(r=radar_df['Model 2'].values,theta=radar_df['Metric'].values,fill='toself',name='Model 2',line_color='#9b59b6'))
            radar_fig.update_layout(polar=dict(radialaxis=dict(visible=True,range=[0, 1])),showlegend=True,title='Model Performance Radar Chart (Higher is Better for All Metrics)',template='plotly_white')
            st.plotly_chart(radar_fig, use_container_width=True)
            st.subheader("Prediction Accuracy Heat Map")
            heatmap_data = pd.DataFrame({'Bag Plus Plant': df['Bag Plus Plant'],'Neural Network Error (%)': df['Error_Percent_Model1'],'Ensemble Algorithm Error (%)': df['Error_Percent_Model2']})
            heatmap_pivot = heatmap_data.set_index('Bag Plus Plant')
            heatmap_fig = px.imshow(heatmap_pivot.T,text_auto='.1f',aspect="auto",color_continuous_scale='RdYlGn_r',title='Prediction Error Heat Map (%)',labels=dict(x="Cement Bag Type", y="Model", color="Error (%)"))
            heatmap_fig.update_layout(height=400, template='plotly_white')
            if len(df) > 5:
                heatmap_fig.update_layout(xaxis_tickangle=-45)
            st.plotly_chart(heatmap_fig, use_container_width=True)
            st.subheader("Error Trend Analysis for High Volume Products")
            top_products = min(5, len(df))
            top_df = df.sort_values('Mar-Actual', ascending=False).head(top_products)
            trend_fig = make_subplots(specs=[[{"secondary_y": True}]])
            trend_fig.add_trace(go.Bar(x=top_df['Bag Plus Plant'], y=top_df['Mar-Actual'], name='Actual Consumption',marker_color='rgba(46, 204, 113, 0.7)'),secondary_y=False,)
            trend_fig.add_trace(go.Scatter(x=top_df['Bag Plus Plant'], y=top_df['Error_Percent_Model1'],mode='lines+markers', name='Neural Network Error (%)',line=dict(color='#3498db', width=2)),secondary_y=True,)
            trend_fig.add_trace(go.Scatter(x=top_df['Bag Plus Plant'], y=top_df['Error_Percent_Model2'], mode='lines+markers', name='Ensemble Error (%)',line=dict(color='#9b59b6', width=2)),secondary_y=True,)
            trend_fig.update_layout(title_text="High Volume Products: Actual Consumption vs Error Percentage",template='plotly_white',barmode='group')
            trend_fig.update_yaxes(title_text="Actual Consumption", secondary_y=False)
            trend_fig.update_yaxes(title_text="Error Percentage (%)", secondary_y=True)
            st.plotly_chart(trend_fig, use_container_width=True)
            st.subheader("Model Stability Analysis")
            model1_std = np.std(df['Error_Percent_Model1'])
            model2_std = np.std(df['Error_Percent_Model2'])
            model1_q75, model1_q25 = np.percentile(df['Error_Percent_Model1'], [75, 25])
            model2_q75, model2_q25 = np.percentile(df['Error_Percent_Model2'], [75, 25])
            model1_iqr = model1_q75 - model1_q25
            model2_iqr = model2_q75 - model2_q25
            stability_data = pd.DataFrame({'Metric': ['Standard Deviation of Errors (%)', 'Interquartile Range (IQR) of Errors (%)','Maximum Error (%)', 'Minimum Error (%)', 'Range of Errors (%)'],'Neural Network': [model1_std, model1_iqr, df['Error_Percent_Model1'].max(), df['Error_Percent_Model1'].min(),df['Error_Percent_Model1'].max() - df['Error_Percent_Model1'].min()],'Ensemble Algorithm': [model2_std, model2_iqr,df['Error_Percent_Model2'].max(), df['Error_Percent_Model2'].min(),df['Error_Percent_Model2'].max() - df['Error_Percent_Model2'].min()]})
            stability_better = []
            for i in range(len(stability_data)):
                if stability_data['Neural Network'].iloc[i] < stability_data['Ensemble Algorithm'].iloc[i]:
                    stability_better.append("Neural Network")
                elif stability_data['Neural Network'].iloc[i] > stability_data['Ensemble Algorithm'].iloc[i]:
                    stability_better.append("Ensemble Algorithm")
                else:
                    stability_better.append("Equal")
            stability_data['Better Model'] = stability_better
            st.write("Model Stability Metrics (Lower is Better):")
            st.dataframe(stability_data.style.apply(lambda x: ['background-color: #d4f1dd' if v == "Neural Network" else 'background-color: #d1e7f0' if v == "Ensemble Algorithm" else '' for v in x], subset=['Better Model']))
            st.subheader("Error Distribution Box Plot")
            box_data = pd.DataFrame({'Neural Network': df['Error_Percent_Model1'],'Ensemble Algorithm': df['Error_Percent_Model2']})
            box_melted = pd.melt(box_data, var_name='Model', value_name='Percentage Error')
            box_fig = px.box(box_melted, x='Model', y='Percentage Error',color='Model',color_discrete_map={'Neural Network': '#3498db', 'Ensemble Algorithm': '#9b59b6'},title='Error Distribution Comparison',points="all")
            box_fig.update_traces(quartilemethod="exclusive")
            box_fig.update_layout(template='plotly_white')
            st.plotly_chart(box_fig, use_container_width=True)
            st.subheader("Product-level Analysis")
            product_analysis = pd.DataFrame({'Bag Plus Plant': df['Bag Plus Plant'],'Actual Consumption': df['Mar-Actual'],'Neural Network Error (%)': df['Error_Percent_Model1'],'Ensemble Error (%)': df['Error_Percent_Model2']})
            better_model_list = []
            for i in range(len(product_analysis)):
                if product_analysis['Neural Network Error (%)'].iloc[i] < product_analysis['Ensemble Error (%)'].iloc[i]:
                    better_model_list.append("Neural Network")
                elif product_analysis['Neural Network Error (%)'].iloc[i] > product_analysis['Ensemble Error (%)'].iloc[i]:
                    better_model_list.append("Ensemble")
                else:
                    better_model_list.append("Equal")
            product_analysis['Better Model'] = better_model_list
            nn_better_count = sum(1 for model in better_model_list if model == "Neural Network")
            ensemble_better_count = sum(1 for model in better_model_list if model == "Ensemble")
            equal_count = sum(1 for model in better_model_list if model == "Equal")
            st.write("Analysis by Product:")
            st.dataframe(product_analysis.style.apply(lambda x: ['background-color: #d4f1dd' if v == "Neural Network" else 'background-color: #d1e7f0' if v == "Ensemble" else '' for v in x], subset=['Better Model']))
            labels = ['Neural Network Better', 'Ensemble Better', 'Equal Performance']
            values = [nn_better_count, ensemble_better_count, equal_count]
            pie_colors = ['#3498db', '#9b59b6', '#95a5a6']
            pie_fig = go.Figure(data=[go.Pie(labels=labels, values=values, hole=.4, marker_colors=pie_colors)])
            pie_fig.update_layout(title_text='Better Model by Product Count')
            st.plotly_chart(pie_fig, use_container_width=True)
            st.subheader("Value-weighted Analysis")
            weighted_performance = pd.DataFrame({'Bag Plus Plant': df['Bag Plus Plant'],'Actual Consumption': df['Mar-Actual'],'Weight (% of Total)': df['Mar-Actual'] / df['Mar-Actual'].sum() * 100,'Neural Network Error (%)': df['Error_Percent_Model1'],'Ensemble Error (%)': df['Error_Percent_Model2'],'Weighted NN Error': df['Error_Percent_Model1'] * df['Mar-Actual'] / df['Mar-Actual'].sum(),'Weighted Ensemble Error': df['Error_Percent_Model2'] * df['Mar-Actual'] / df['Mar-Actual'].sum()})
            total_weighted_nn = weighted_performance['Weighted NN Error'].sum()
            total_weighted_ensemble = weighted_performance['Weighted Ensemble Error'].sum()
            st.write("Value-weighted Error Analysis (Higher volume products have more weight):")
            st.dataframe(weighted_performance[['Bag Plus Plant', 'Actual Consumption', 'Weight (% of Total)','Neural Network Error (%)', 'Ensemble Error (%)']].style.background_gradient(cmap='RdYlGn_r', subset=['Neural Network Error (%)', 'Ensemble Error (%)']))
            weighted_col1, weighted_col2 = st.columns(2)
            with weighted_col1:
                st.metric("Total Weighted Error - Neural Network", f"{total_weighted_nn:.2f}%")
            with weighted_col2:
                st.metric("Total Weighted Error - Ensemble Algorithm", f"{total_weighted_ensemble:.2f}%")
            winner_value_weighted = "Neural Network" if total_weighted_nn < total_weighted_ensemble else "Ensemble Algorithm"
            st.info(f"Value-weighted Winner: **{winner_value_weighted}**")
            st.header("Download Reports")
            st.subheader("Download Analysis Results as Excel")
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, sheet_name='Original Data', index=False)
                comparison_df.to_excel(writer, sheet_name='Metrics Comparison', index=False)
                stability_data.to_excel(writer, sheet_name='Stability Analysis', index=False)
                product_analysis.to_excel(writer, sheet_name='Product Analysis', index=False)
                weighted_performance.to_excel(writer, sheet_name='Value-weighted Analysis', index=False)
                workbook = writer.book
                worksheet = writer.sheets['Metrics Comparison']
                better_format = workbook.add_format({'bg_color': '#d4f1dd'})
                worse_format = workbook.add_format({'bg_color': '#f8d7da'})
                for row_num, model in enumerate(better_model, start=1):
                    if model == 'Model 1':
                        worksheet.conditional_format(row_num, 1, row_num, 1, {'type': 'no_blanks', 'format': better_format})
                        worksheet.conditional_format(row_num, 2, row_num, 2, {'type': 'no_blanks', 'format': worse_format})
                    elif model == 'Model 2':
                        worksheet.conditional_format(row_num, 1, row_num, 1, {'type': 'no_blanks', 'format': worse_format})
                        worksheet.conditional_format(row_num, 2, row_num, 2, {'type': 'no_blanks', 'format': better_format})
            excel_data = output.getvalue()
            st.download_button(label="Download Analysis as Excel",data=excel_data,file_name="cement_model_comparison_analysis.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        else:
            st.error(f"Required columns not found. Please ensure your Excel file has these columns: {', '.join(required_columns)}")
    except Exception as e:
        st.error(f"Error processing the file: {str(e)}")
else:
    st.info("Please upload an Excel file with the following columns: 'Bag Plus Plant', 'Mar-Actual', 'Mar Pred1', 'Mar Pred2'")
    sample_df = pd.DataFrame({'Bag Plus Plant': ['Cement Type A - Plant 1', 'Cement Type B - Plant 2', 'Cement Type C - Plant 1'],'Mar-Actual': [1500, 2000, 1200],'Mar Pred1': [1450, 2100, 1250],'Mar Pred2': [1530, 1950, 1180]})
    st.write("Sample data structure:")
    st.dataframe(sample_df)
st.markdown("""<div style="text-align: center; margin-top: 40px; padding: 20px; background-color: #f8f9fa; border-radius: 5px;"><p style="color: #7f8c8d;">Cement Consumption Model Comparison Dashboard</p></div>""", unsafe_allow_html=True)
