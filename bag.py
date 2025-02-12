import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import warnings
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
warnings.filterwarnings('ignore')

def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def select_best_model(data):
    """
    Optimized model selection using only Holt-Winters
    Returns the best model, forecast, and model description
    """
    if len(data) < 12:
        return None, None, "Forecasting not possible due to insufficient data (minimum 12 months required)"
    
    try:
        # Prepare data
        data = data.astype(float)
        
        # Use only Holt-Winters model for faster processing
        hw_model = ExponentialSmoothing(
            data, 
            seasonal_periods=12, 
            trend='add', 
            seasonal='add',
            initialization_method='estimated'
        ).fit(optimized=True)
        
        forecast = hw_model.forecast(1)[0]
        return hw_model, forecast, "Holt-Winters Exponential Smoothing"

    except Exception as e:
        return None, None, f"Error in model fitting: {str(e)}"

def generate_pdf_report(predictions_data, filename="cement_plant_report.pdf"):
    """Generate PDF report with predictions table"""
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    # Create story for PDF content
    story = []
    
    # Add title
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=24,
        spaceAfter=30
    )
    story.append(Paragraph("Cement Plant Demand Forecast Report", title_style))
    
    # Add date
    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontSize=12,
        spaceAfter=30
    )
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%d %B %Y')}", date_style))
    
    # Create table data
    table_data = [['Plant Name', 'Bag Type', 'Predicted Demand (Feb)', 'Actual Demand (Till 9th)', 'Status']]
    for row in predictions_data:
        table_data.append([
            row['plant_name'],
            row['bag_name'],
            f"{row['predicted_demand']:,.2f}",
            f"{row['actual_demand']:,.2f}",
            'Within Range' if row['status'] == 'green' else 'Outside Range'
        ])
    
    # Create table
    table = Table(table_data, colWidths=[2*inch, 2*inch, 1.5*inch, 1.5*inch, 1*inch])
    
    # Add style to table
    style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (2, 1), (3, -1), 'RIGHT'),
    ])
    
    # Add row colors based on status
    for i in range(1, len(table_data)):
        if predictions_data[i-1]['status'] == 'red':
            style.add('BACKGROUND', (4, i), (4, i), colors.pink)
        else:
            style.add('BACKGROUND', (4, i), (4, i), colors.lightgreen)
    
    table.setStyle(style)
    story.append(table)
    
    # Build PDF
    doc.build(story)

def main():
    st.set_page_config("Cement Plant Bag Usage Analysis", layout="wide")
    st.title("Cement Plant Bag Usage Analysis")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])
        actual_demand_file = st.file_uploader("Upload February Actual Demand (till 9th) Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None and actual_demand_file is not None:
        # Read the Excel files
        df = pd.read_excel(uploaded_file)
        actual_df = pd.read_excel(actual_demand_file)
        
        # Remove the first column
        df = df.iloc[:, 1:]
        
        # Get unique plant names
        unique_plants = sorted(df['Cement Plant Sname'].unique())
        
        # Store predictions for PDF report
        predictions_data = []
        
        for plant in unique_plants:
            plant_bags = df[df['Cement Plant Sname'] == plant]['MAKTX'].unique()
            
            for bag in plant_bags:
                selected_data = df[(df['Cement Plant Sname'] == plant) & 
                                 (df['MAKTX'] == bag)]
                
                if not selected_data.empty:
                    month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
                    
                    # Create time series data
                    all_usage_data = []
                    for month in month_columns:
                        date = pd.to_datetime(month)
                        usage = selected_data[month].iloc[0]
                        all_usage_data.append({
                            'Date': date,
                            'Usage': usage
                        })
                    
                    all_data_df = pd.DataFrame(all_usage_data)
                    all_data_df = all_data_df.sort_values('Date')
                    
                    # Prepare time series data
                    ts_data = all_data_df.set_index('Date')['Usage']
                    
                    # Get forecast
                    model, forecast, _ = select_best_model(ts_data)
                    
                    if forecast is not None:
                        # Calculate predicted demand till 9th Feb
                        predicted_9th = (9/28) * forecast
                        
                        # Get actual demand
                        actual_demand = actual_df[
                            (actual_df['Cement Plant Sname'] == plant) & 
                            (actual_df['MAKTX'] == bag)
                        ]['Actual_Demand'].iloc[0]
                        
                        # Calculate percentage difference
                        diff_percentage = abs((actual_demand - predicted_9th) / predicted_9th * 100)
                        status = 'red' if diff_percentage > 10 else 'green'
                        
                        # Store prediction data
                        predictions_data.append({
                            'plant_name': plant,
                            'bag_name': bag,
                            'predicted_demand': forecast,
                            'actual_demand': actual_demand,
                            'status': status
                        })
        
        # Generate PDF report
        if predictions_data:
            generate_pdf_report(predictions_data)
            st.success("PDF report has been generated successfully!")
            
            # Display results in app
            st.subheader("Demand Comparison")
            for pred in predictions_data:
                color = "red" if pred['status'] == 'red' else "green"
                st.markdown(
                    f"""
                    <div style='padding: 10px; border-radius: 5px; background-color: {color}20; border: 1px solid {color}'>
                        <h3 style='color: {color}'>{pred['plant_name']} - {pred['bag_name']}</h3>
                        <p>Predicted February Demand: {pred['predicted_demand']:,.2f}</p>
                        <p>Actual Demand (Till 9th): {pred['actual_demand']:,.2f}</p>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

if __name__ == '__main__':
    main()
