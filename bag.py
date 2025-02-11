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

def select_best_model(data):
    """
    Optimized model selection using only Holt-Winters
    Returns the model, forecast, and model description
    """
    if len(data) < 12:
        return None, None, "Forecasting not possible due to insufficient data (minimum 12 months required)"
    
    try:
        # Prepare data
        data = data.astype(float)
        
        # Use only Holt-Winters model for faster processing
        hw_model = ExponentialSmoothing(data, 
                                      seasonal_periods=12, 
                                      trend='add', 
                                      seasonal='add').fit()
        
        forecast = hw_model.forecast(1)[0]
        return hw_model, forecast, "Holt-Winters Exponential Smoothing"

    except Exception as e:
        return None, None, f"Error in model fitting: {str(e)}"

def generate_pdf_report(predictions_data, filename="demand_forecast_report.pdf"):
    """Generate PDF report with forecasting results"""
    doc = SimpleDocTemplate(filename, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=72)
    elements = []
    
    # Styles
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Title'],
        fontSize=24,
        spaceAfter=30
    )
    subtitle_style = ParagraphStyle(
        'CustomSubtitle',
        parent=styles['Normal'],
        fontSize=14,
        textColor=colors.grey,
        spaceAfter=20
    )
    
    # Add title and subtitle
    title = Paragraph("Cement Plant Demand Forecast Report", title_style)
    subtitle = Paragraph(f"Generated on {datetime.now().strftime('%d %B %Y')}", subtitle_style)
    elements.extend([title, subtitle])
    
    # Create table data
    table_data = [['Plant Name', 'Bag Type', 'Predicted Demand (Feb)', 'Status']]
    for row in predictions_data:
        status_color = colors.green if row['status'] == 'Within Range' else colors.red
        table_data.append([
            row['plant'],
            row['bag'],
            f"{row['prediction']:,.2f}",
            row['status']
        ])
    
    # Create table with styling
    table = Table(table_data, colWidths=[2.5*inch, 2*inch, 1.5*inch, 1*inch])
    table.setStyle(TableStyle([
        # Header styling
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#333333')),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('TOPPADDING', (0, 0), (-1, 0), 12),
        
        # Body styling
        ('BACKGROUND', (0, 1), (-1, -1), colors.white),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('ALIGN', (2, 1), (2, -1), 'RIGHT'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        
        # Alternating row colors
        ('ROWBACKGROUNDS', (0, 1), (-1, -1), [colors.white, colors.HexColor('#f8f9fa')]),
        
        # Cell padding
        ('TOPPADDING', (0, 1), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
    ]))
    
    elements.append(table)
    
    # Add footer
    elements.append(Spacer(1, 30))
    footer_text = """Note: Status indicates whether the actual demand (Feb 1-9) is within 10% of predicted demand.
                    'Within Range' (Green) indicates variance ≤ 10%, 'Outside Range' (Red) indicates variance > 10%."""
    footer = Paragraph(footer_text, styles['Italic'])
    elements.append(footer)
    
    # Build PDF
    doc.build(elements)
    return filename

def create_usage_plot(plot_data, selected_plant, selected_bag, forecast_info=None):
    """Create plotly figure for usage visualization"""
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
    if forecast_info and 'prediction' in forecast_info:
        fig.add_trace(go.Scatter(
            x=['Feb 2025'],
            y=[forecast_info['prediction']],
            name='Forecasted (Feb)',
            mode='markers',
            marker=dict(
                color='red' if forecast_info['status'] == 'Outside Range' else 'green',
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
    
    # Update layout
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
    
    return fig

def main():
    st.set_page_config("Cement Plant Bag Usage Analysis", layout="wide")
    st.title("Cement Plant Bag Usage Analysis")
    
    # Create two columns
    col1, col2 = st.columns([2, 1])
    
    with col1:
        uploaded_file = st.file_uploader("Upload your Excel files", type=['xlsx', 'xls'])
        actual_data_file = st.file_uploader("Upload February 1-9 actual data (Excel)", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Read the Excel files
        df = pd.read_excel(uploaded_file)
        actual_feb_data = pd.read_excel(actual_data_file) if actual_data_file else None
        
        # Remove the first column
        df = df.iloc[:, 1:]
        
        # Get unique plant names
        unique_plants = sorted(df['Cement Plant Sname'].unique())
        
        # Store predictions for PDF report
        predictions_data = []
        
        with st.spinner('Generating forecasts for all plants...'):
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
                        ts_data = all_data_df.set_index('Date')['Usage']
                        
                        # Get forecast
                        model, forecast, _ = select_best_model(ts_data)
                        
                        if forecast is not None:
                            # Calculate predicted demand for Feb 1-9
                            predicted_9_days = (forecast * 9) / 28
                            
                            # Get actual demand if available
                            actual_9_days = None
                            if actual_feb_data is not None:
                                actual_data = actual_feb_data[
                                    (actual_feb_data['Cement Plant Sname'] == plant) & 
                                    (actual_feb_data['MAKTX'] == bag)
                                ]
                                if not actual_data.empty:
                                    actual_9_days = actual_data['Usage'].iloc[0]
                            
                            # Determine status
                            status = 'N/A'
                            if actual_9_days is not None:
                                diff_percentage = abs((actual_9_days - predicted_9_days) / predicted_9_days * 100)
                                status = 'Within Range' if diff_percentage <= 10 else 'Outside Range'
                            
                            predictions_data.append({
                                'plant': plant,
                                'bag': bag,
                                'prediction': forecast,
                                'status': status,
                                'actual_9_days': actual_9_days,
                                'predicted_9_days': predicted_9_days
                            })
        
        # Generate PDF report
        if predictions_data:
            pdf_file = generate_pdf_report(predictions_data)
            
            with col1:
                st.success("Forecasts generated successfully!")
                with open(pdf_file, "rb") as file:
                    st.download_button(
                        label="Download Forecast Report (PDF)",
                        data=file,
                        file_name="demand_forecast_report.pdf",
                        mime="application/pdf"
                    )
        
        # Display selected plant analysis
        with col1:
            selected_plant = st.selectbox('Select Cement Plant:', unique_plants)
            plant_bags = df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique()
            selected_bag = st.selectbox('Select Bag:', sorted(plant_bags))
            
            # Get selected data
            selected_data = df[(df['Cement Plant Sname'] == selected_plant) & 
                             (df['MAKTX'] == selected_bag)]
            
            if not selected_data.empty:
                # Get forecast info for selected plant/bag
                selected_prediction = next(
                    (p for p in predictions_data 
                     if p['plant'] == selected_plant and p['bag'] == selected_bag), 
                    None
                )
                
                # Display forecast comparison
                if selected_prediction:
                    status_color = "#28a745" if selected_prediction['status'] == 'Within Range' else "#dc3545"
                    
                    # Create three columns for metrics
                    m1, m2, m3 = st.columns(3)
                    
                    with m1:
                        st.metric(
                            "Predicted Feb Demand",
                            f"{selected_prediction['prediction']:,.2f}"
                        )
                    
                    with m2:
                        if selected_prediction['actual_9_days'] is not None:
                            st.metric(
                                "Actual (Feb 1-9)",
                                f"{selected_prediction['actual_9_days']:,.2f}",
                                f"{((selected_prediction['actual_9_days'] - selected_prediction['predicted_9_days']) / selected_prediction['predicted_9_days'] * 100):,.1f}%"
                            )
                    
                    with m3:
                        st.markdown(
                            f"""
                            <div style="padding: 1rem; border-radius: 0.5rem; background-color: {status_color}; color: white; text-align: center;">
                                <h3 style="margin: 0;">Status: {selected_prediction['status']}</h3>
                            </div>
                            """,
                            unsafe_allow_html=True
                        )
                
                # Create monthly data for plotting
                month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
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
                
                # Create and display the plot
                fig = create_usage_plot(plot_data, selected_plant, selected_bag, selected_prediction)
                st.plotly_chart(fig, use_container_width=True)
                
                # Display the complete historical data
                st.subheader("Complete Historical Data")
                display_df = all_data_df[['Month', 'Usage']]
                st.dataframe(
                    display_df.style.format({'Usage': '{:,.2f}'}).apply(
                        lambda x: ['background-color: #f8f9fa' if i % 2 else '' 
                                 for i in range(len(x))
                        ], axis=0
                    )
                )

if __name__ == '__main__':
    main()
