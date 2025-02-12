import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import numpy as np
from statsmodels.tsa.holtwinters import ExponentialSmoothing
import warnings
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
warnings.filterwarnings('ignore')

def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def optimize_model_selection(data):
    """
    Optimized model selection focusing on Holt-Winters with simplified parameters
    Returns the model, forecast, and model description
    """
    if len(data) < 12:
        return None, None, "Insufficient data (minimum 12 months required)"
    
    try:
        # Use only Holt-Winters with optimized parameters
        hw_model = ExponentialSmoothing(
            data,
            seasonal_periods=12,
            trend='add',
            seasonal='add',
            initialization_method='estimated'
        ).fit(optimized=True, use_boxcox=True)
        
        forecast = hw_model.forecast(1)[0]
        return hw_model, forecast, "Holt-Winters Exponential Smoothing"

    except Exception as e:
        return None, None, f"Error in model fitting: {str(e)}"

def generate_pdf_report(data, filename="cement_plant_report.pdf"):
    """Generate PDF report with forecasting results"""
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
        spaceAfter=30,
        alignment=1  # Center alignment
    )
    story.append(Paragraph("Cement Plant Demand Forecast Report", title_style))
    story.append(Spacer(1, 20))
    
    # Add date
    date_style = ParagraphStyle(
        'DateStyle',
        parent=styles['Normal'],
        fontSize=12,
        alignment=1
    )
    story.append(Paragraph(f"Generated on: {datetime.now().strftime('%B %d, %Y')}", date_style))
    story.append(Spacer(1, 30))
    
    # Create table data
    table_data = [['Plant Name', 'Bag Type', 'Forecasted Demand (Feb 2025)']]
    for row in data:
        table_data.append([
            row['plant'],
            row['bag'],
            f"{row['forecast']:,.2f}"
        ])
    
    # Create table
    table = Table(table_data, colWidths=[2.5*inch, 2.5*inch, 2*inch])
    table.setStyle(TableStyle([
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
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('PADDING', (0, 0), (-1, -1), 8),
    ]))
    
    story.append(table)
    doc.build(story)
    return filename

def main():
    st.title("Cement Plant Bag Usage Analysis")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df = df.iloc[:, 1:]
        unique_plants = sorted(df['Cement Plant Sname'].unique())
        
        # Store forecasting results for PDF
        forecast_results = []
        
        for plant in unique_plants:
            plant_bags = df[df['Cement Plant Sname'] == plant]['MAKTX'].unique()
            
            for bag in plant_bags:
                selected_data = df[(df['Cement Plant Sname'] == plant) & 
                                 (df['MAKTX'] == bag)]
                
                if not selected_data.empty:
                    month_columns = [col for col in df.columns 
                                   if col not in ['Cement Plant Sname', 'MAKTX']]
                    
                    # Prepare time series data
                    all_usage_data = []
                    for month in month_columns:
                        date = pd.to_datetime(month)
                        usage = selected_data[month].iloc[0]
                        all_usage_data.append({
                            'Date': date,
                            'Usage': usage
                        })
                    
                    ts_data = pd.DataFrame(all_usage_data)
                    ts_data = ts_data.set_index('Date')['Usage']
                    
                    # Get forecast
                    model, forecast, _ = optimize_model_selection(ts_data)
                    
                    if forecast is not None:
                        forecast_results.append({
                            'plant': plant,
                            'bag': bag,
                            'forecast': forecast
                        })
        
        if forecast_results:
            # Generate PDF report
            with st.spinner('Generating PDF report...'):
                pdf_file = generate_pdf_report(forecast_results)
                
            # Provide download button for PDF
            with open(pdf_file, "rb") as f:
                st.download_button(
                    label="Download PDF Report",
                    data=f,
                    file_name="cement_plant_report.pdf",
                    mime="application/pdf"
                )

if __name__ == '__main__':
    main()
