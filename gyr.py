import streamlit as st
import pandas as pd
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from datetime import datetime
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
import base64
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
# Sidebar navigation
with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None

# Load Lottie animations
lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")
import numpy as np
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from io import BytesIO
import numpy as np
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from io import BytesIO
import numpy as np
from datetime import datetime
from reportlab.lib.pagesizes import letter, A4
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from io import BytesIO
import numpy as np
from datetime import datetime
import matplotlib.pyplot as plt

def create_pdf_report(region, df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    styles = getSampleStyleSheet()

    def add_page_number(canvas, doc):
        page_num = canvas.getPageNumber()
        text = f"Page {page_num}"
        canvas.saveState()
        canvas.setFillColor(colors.grey)
        canvas.setFont("Helvetica", 8)
        canvas.drawRightString(width - 30, 30, text)
        canvas.restoreState()

    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.savefig(img_buffer, format="png", dpi=300, bbox_inches="tight")
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.darkblue),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 8),
            ('TOPPADDING', (0, 1), (-1, -1), 3),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)

    def add_header():
        c.setFillColor(colors.darkblue)
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 22)
        c.drawString(30, height - 35, f"GYR Analysis Report: {region}")
        c.setFont("Helvetica", 12)
        c.drawString(30, height - 48, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    def add_front_page():
        c.setFillColor(colors.lightblue)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColor(colors.darkblue)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "GYR Analysis Report")
        c.setFont("Helvetica", 24)
        c.drawCentredString(width / 2, height - 250, f"Region: {region}")
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        
        # Add a professional logo or image
        logo_path = ""  # Replace with your logo path
        c.drawImage(logo_path, width/2 - 50, height - 150, width=100, height=100)
        
        c.showPage()

    def add_executive_summary(df):
        c.setFont("Helvetica-Bold", 18)
        c.drawString(50, height - 100, "Executive Summary")
        
        summary_style = ParagraphStyle('Summary', fontName='Helvetica', fontSize=10, leading=14, spaceAfter=10)
        
        summary_text = f"""
        This report provides a comprehensive analysis of the GYR (Green, Yellow, Red) metrics for the {region} region. 
        Key findings include:
        
        1. Overall Performance: The region shows a {get_trend(df['Overall EBITDA'])} trend in overall EBITDA.
        2. Green Products: {get_product_summary(df, 'Green')}
        3. Yellow Products: {get_product_summary(df, 'Yellow')}
        4. Red Products: {get_product_summary(df, 'Red')}
        5. Recommendations: {get_recommendations(df)}
        """
        
        p = Paragraph(summary_text, summary_style)
        p.wrapOn(c, width - 100, height)
        p.drawOn(c, 50, height - 350)
        
        c.showPage()

    def get_trend(series):
        if series.iloc[-1] > series.iloc[0]:
            return "positive"
        elif series.iloc[-1] < series.iloc[0]:
            return "negative"
        else:
            return "stable"

    def get_product_summary(df, color):
        share = df[f'{color}'].mean() / (df['Green'] + df['Yellow'] + df['Red']).mean() * 100
        trend = get_trend(df[f'{color} EBITDA'])
        return f"Represents {share:.1f}% of products with a {trend} EBITDA trend"

    def get_recommendations(df):
        green_share = df['Green'].mean() / (df['Green'] + df['Yellow'] + df['Red']).mean()
        if green_share < 0.3:
            return "Focus on increasing the share of Green products to improve overall EBITDA"
        elif df['Red EBITDA'].mean() < df['Yellow EBITDA'].mean() * 0.8:
            return "Implement strategies to improve the performance of Red products"
        else:
            return "Maintain current strategy while monitoring market trends"

    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(50, height - 100, "Appendix")
        
        appendix_style = ParagraphStyle('Appendix', fontName='Helvetica', fontSize=10, leading=14, spaceAfter=10)
        
        appendix_text = """
        1. Methodology:
           - Data collection: Monthly sales and EBITDA data for Green, Yellow, and Red products
           - Analysis: Trend analysis, share distribution, and comparative performance evaluation
        
        2. Key Metric Definitions:
           - Green Products: High-performance products with the best EBITDA margins
           - Yellow Products: Mid-range products with moderate EBITDA margins
           - Red Products: Products requiring improvement or potential phase-out
        
        3. Limitations:
           - This analysis is based on historical data and may not predict future market changes
           - External factors such as economic conditions are not accounted for in this report
        
        4. Further Analysis Recommendations:
           - Conduct customer segmentation analysis to identify target markets for each product category
           - Perform a detailed cost analysis to identify opportunities for improving Red product performance
           - Investigate successful strategies in regions with high Green product share for potential replication
        """
        
        p = Paragraph(appendix_text, appendix_style)
        p.wrapOn(c, width - 100, height)
        p.drawOn(c, 50, height - 500)

    add_front_page()
    add_executive_summary(df)
    
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header()
                    cols = ['Green EBITDA', 'Yellow EBITDA', 'Red EBITDA']
                    overall_col = 'Overall EBITDA'

                    total_quantity = filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red']
                    filtered_df[overall_col] = (
                        (filtered_df['Green'] * filtered_df['Green EBITDA'] +
                         filtered_df['Yellow'] * filtered_df['Yellow EBITDA'] + 
                         filtered_df['Red'] * filtered_df['Red EBITDA']) / total_quantity
                    )

                    filtered_df['Average Green Share'] = filtered_df['Green'] / total_quantity
                    filtered_df['Average Yellow Share'] = filtered_df['Yellow'] / total_quantity
                    filtered_df['Average Red Share'] = filtered_df['Red'] / total_quantity
                    
                    def adjust_shares(row):
                        green, yellow, red = row['Average Green Share'], row['Average Yellow Share'], row['Average Red Share']
                        
                        if green == 1 or yellow == 1 or red == 1:
                            return green, yellow, red
                        elif green == 0 and yellow == 0:
                            return green, yellow, red
                        elif green == 0:
                            yellow = min(yellow + 0.05, 1)
                            red = max(1 - yellow, 0)
                        elif yellow == 0:
                            green = min(green + 0.05, 1)
                            red = max(1 - green, 0)
                        else:
                            green = min(green + 0.05, 1)
                            yellow = min(yellow + 0.025, 1 - green)
                            red = max(1 - green - yellow, 0)
                        
                        return green, yellow, red

                    filtered_df['Adjusted Green Share'], filtered_df['Adjusted Yellow Share'], filtered_df['Adjusted Red Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Green Share'] * filtered_df['Green EBITDA'] +
                        filtered_df['Adjusted Yellow Share'] * filtered_df['Yellow EBITDA'] +
                        filtered_df['Adjusted Red Share'] * filtered_df['Red EBITDA']
                    )

                    filtered_df['G-R Difference'] = filtered_df['Green EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['G-Y Difference'] = filtered_df['Green EBITDA'] - filtered_df['Yellow EBITDA']
                    filtered_df['Y-R Difference'] = filtered_df['Yellow EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 12), gridspec_kw={'height_ratios': [3, 1]})
                    
                    ax1.plot(filtered_df['Month'], filtered_df['Green EBITDA'], 'g-', label='Green EBITDA')
                    ax1.plot(filtered_df['Month'], filtered_df['Yellow EBITDA'], 'y-', label='Yellow EBITDA')
                    ax1.plot(filtered_df['Month'], filtered_df['Red EBITDA'], 'r-', label='Red EBITDA')
                    ax1.plot(filtered_df['Month'], filtered_df[overall_col], 'b--', label='Overall EBITDA')
                    ax1.plot(filtered_df['Month'], filtered_df['Imaginary EBITDA'], 'm:', label='Imaginary EBITDA')
                    
                    ax1.set_title(f"EBITDA Analysis for {brand} (Type: {product_type}) in {region} ({region_subset})")
                    ax1.set_xlabel('Month')
                    ax1.set_ylabel('EBITDA (Rs./MT)')
                    ax1.legend()
                    ax1.grid(True, linestyle='--', alpha=0.7)
                    
                    ax2.plot(filtered_df['Month'], filtered_df['I-O Difference'], 'k-', label='I-O Difference')
                    ax2.axhline(y=filtered_df['I-O Difference'].mean(), color='r', linestyle='--', label=f'Mean: {filtered_df["I-O Difference"].mean():.2f}')
                    
                    ax2.set_xlabel('Month')
                    ax2.set_ylabel('I-O Difference (Rs./MT)')
                    ax2.legend()
                    ax2.grid(True, linestyle='--', alpha=0.7)
                    
                    plt.tight_layout()
                    
                    draw_graph(fig, 50, height - 450, 500, 400)
                    plt.close(fig)

                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(50, height - 470, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Green','Yellow','Red'] + cols + [overall_col, 'Imaginary EBITDA']].describe().round(2)
                    table_data = [['Metric'] + list(desc_stats.columns)] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 480, [40, 40, 40, 40] + [70] * (len(desc_stats.columns) - 4))
                    
                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(50, height - 650, "Average Share Distribution")
                    
                    fig, ax = plt.subplots(figsize=(6, 6))
                    average_shares = filtered_df[['Average Green Share', 'Average Yellow Share', 'Average Red Share']].mean()
                    ax.pie(average_shares.values, labels=average_shares.index, autopct='%1.1f%%', startangle=90, colors=['green', 'yellow', 'red'])
                    ax.axis('equal')
                    plt.title("Average Share Distribution")
                    draw_graph(fig, 80, height - 850, 200, 200)
                    plt.close(fig)

                    c.setFont("Helvetica-Bold", 12)
                    c.drawString(330, height - 650, "Monthly Share Distribution")
                    
                    share_data = [['Month', 'Green', 'Yellow', 'Red']]
                    for _, row in filtered_df[['Month', 'Green', 'Yellow', 'Red', 'Average Green Share', 'Average Yellow Share', 'Average Red Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Green']:.0f} ({row['Average Green Share']:.2%})",
                            f"{row['Yellow']:.0f} ({row['Average Yellow Share']:.2%})",
                            f"{row['Red']:.0f} ({row['Average Red Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 670, [40, 60, 60, 60])
                    
                    # Add key insights
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(50, height - 880, "Key Insights")
                    
                    insights_style = ParagraphStyle('Insights', fontName='Helvetica', fontSize=10, leading=14, spaceAfter=10)
                    
                    insights_text = f"""
                    1. EBITDA Trends:
                       - Green products show a {get_trend(filtered_df['Green EBITDA'])} trend
                       - Yellow products show a {get_trend(filtered_df['Yellow EBITDA'])} trend
                       - Red products show a {get_trend(filtered_df['Red EBITDA'])} trend
                    
                    2. Share Distribution:
                       - Green products account for {average_shares['Average Green Share']*100:.1f}% of the total
                       - Yellow products account for {average_shares['Average Yellow Share']*100:.1f}% of the total
                       - Red products account for {average_shares['Average Red Share']*100:.1f}% of the total
                    
                    3. Imaginary vs Actual EBITDA:
                       - Average difference: {filtered_df['I-O Difference'].mean():.2f} Rs./MT
                       - This suggests potential for improvement by optimizing product mix
                    
                    4. Recommendations:
                       {get_recommendations(filtered_df)}
                    """
                    
                    p = Paragraph(insights_text, insights_style)
                    p.wrapOn(c, width - 100, height)
                    p.drawOn(c, 50, height - 1050)
                    
                    c.showPage()
                    
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer

# Helper functions (if not already defined)
def get_trend(series):
    if series.iloc[-1] > series.iloc[0]:
        return "increasing"
    elif series.iloc[-1] < series.iloc[0]:
        return "decreasing"
    else:
        return "stable"


if selected == "Home":
    st.title("📊 Advanced GYR Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
        if uploaded_file is not None:
            st.session_state.uploaded_file = uploaded_file
            st.success("File successfully uploaded! Please go to the Analysis page to view results.")

    with col2:
        if lottie_upload:
            st_lottie(lottie_upload, height=150, key="upload")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/4503/4503700.png", width=150)
    st.markdown("</div>", unsafe_allow_html=True)
elif selected == "Analysis":
    st.title("📈 Data Analysis Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)

        # Create sidebar for user inputs
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download button for combined report
        if st.sidebar.button(f"Download Combined Report for {region}"):
            pdf_buffer = create_pdf_report(region, df)
            pdf_bytes = pdf_buffer.getvalue()
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}.pdf">Download PDF Report</a>'
            st.sidebar.markdown(href, unsafe_allow_html=True)

        # Add unique keys to each selectbox
        brand = st.sidebar.selectbox("Select Brand", options=df['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df['Region subsets'].unique(), key="region_subset_select")
        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type

        green_share = st.sidebar.slider("Adjust Green Share (%)", 0, 99, 50, key="green_share_slider")
        yellow_share = st.sidebar.slider("Adjust Yellow Share (%)", 0, 100-green_share, 0, key="yellow_share_slider")
        red_share = 100 - green_share - yellow_share
        st.sidebar.text(f"Red Share: {red_share}%")
        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Green NSR', 'Yellow NSR', 'Red NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Green Contribution', 'Yellow Contribution','Red Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Green EBITDA', 'Yellow EBITDA','Red EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Green'] * filtered_df[cols[0]] +
                                        filtered_df['Yellow'] * filtered_df[cols[1]] + filtered_df['Red']*filtered_df[cols[2]]) / (
                                            filtered_df['Green'] + filtered_df['Yellow']+filtered_df['Red'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - (green_share+yellow_share)/100) * filtered_df[cols[2]] +
                                          (green_share/100) * filtered_df[cols[0]] + (yellow_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['G-Y Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            filtered_df['G-R Difference'] = filtered_df[cols[0]] - filtered_df[cols[2]]
            filtered_df['Y-R Difference'] = filtered_df[cols[1]] - filtered_df[cols[2]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            
            if cols[0] in cols:
                  fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[0]],
                                         mode='lines+markers', name=cols[0],line_color="green"))
            if cols[1] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[1]],
                                         mode='lines+markers', name=cols[1],line_color="yellow"))
            if cols[2] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[2]],
                                         mode='lines+markers', name=cols[2],line_color="red"))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({green_share}% Green & {yellow_share}% Yellow)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(G-Y: {diff:.2f})<br>(G-R: {i_diff:.2f})<br>(Y-R: {j_diff:.2f})<br>(I-O: {k_diff:.2f})" for month, diff, i_diff, j_diff, k_diff in 
                        zip(filtered_df['Month'], filtered_df['G-Y Difference'], filtered_df['G-R Difference'], filtered_df['Y-R Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (G-Y: Green - Red,G-R: Green - Red,Y-R: Yellow - Red, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            st.subheader("Descriptive Statistics")
            desc_stats = filtered_df[cols + [overall_col, imaginary_col]].describe()
            st.dataframe(desc_stats.style.format("{:.2f}").background_gradient(cmap='Blues'), use_container_width=True)
                    
                    # Display share of Green, Yellow, and Red Products
            st.subheader("Share of Green, Yellow, and Red Products")
            total_quantity = filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red']
            green_share = (filtered_df['Green'] / total_quantity * 100).round(2)
            yellow_share = (filtered_df['Yellow'] / total_quantity * 100).round(2)
            red_share = (filtered_df['Red'] / total_quantity * 100).round(2)
                    
            share_df = pd.DataFrame({
                        'Month': filtered_df['Month'],
                        'Green Share (%)': green_share,
                        'Yellow Share (%)': yellow_share,
                        'Red Share (%)': red_share
                    })
                    
            fig_pie = px.pie(share_df, values=[green_share.mean(), yellow_share.mean(), red_share.mean()], 
                                     names=['Green', 'Yellow', 'Red'], title='Average Share Distribution',color=["G","Y","R"],color_discrete_map={"G":"green","Y":"yellow","R":"red"},hole=0.5)
            st.plotly_chart(fig_pie, use_container_width=True)
                    
            st.dataframe(share_df.set_index('Month').style.format("{:.2f}").background_gradient(cmap='RdYlGn'), use_container_width=True)
        
        
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)

elif selected == "About":
    st.title("About the GYR Analysis App")
    st.markdown("""
    This advanced data analysis application is designed to provide insightful visualizations and statistics for your GYR (Green, Yellow, Red) data. 
    
    Key features include:
    - Interactive data filtering
    - Multiple analysis types (NSR, Contribution, EBITDA)
    - Dynamic visualizations with Plotly
    - Descriptive statistics and share analysis
    - Customizable Green and Yellow share adjustments
    """)
   
