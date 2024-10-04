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
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER
from io import BytesIO
from datetime import datetime
from reportlab.graphics import renderPDF
import random
from reportlab.lib.units import inch
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
def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 24)
        header_text = f"GYR Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "Segment Mix Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the Segment Mix Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Trade
            [random.randint(1500, 2500) for _ in range(12)],  # Non-Trade
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.blue
        lc.lines[2].strokeColor = colors.pink
        lc.lines[3].strokeColor = colors.brown

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Trade EBITDA'),
            (colors.blue, 'Non-Trade EBITDA'),
            (colors.crimson, 'Overall EBITDA'),
            (colors.brown, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Green, Yellow, and Red EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If both (Trade,Non-Trade) are present: Trade +5%, Non-Trade -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Trade and Non-Trade Channel for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Trade Channel specifically for PPC, which typically have higher EBIDTA.",
            "Analyze factors contributing to higher EBIDTA in Trade Channel,and apply insights to Non-Trade.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all channels.",
            "Invest in product innovation to expand Trade Channel offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for Trade and Non-Trade channel ceteris paribus.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across GYR,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
                    filtered_df[overall_col] = (
                        (filtered_df['Trade'] * filtered_df['Trade EBITDA'] +
                         filtered_df['Non-Trade'] * filtered_df['Non-Trade EBITDA'])/ total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Trade Share'] = filtered_df['Trade'] / total_quantity
                    filtered_df['Average Non-Trade Share'] = filtered_df['Non-Trade'] / total_quantity
                    
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        trade = row['Average Trade Share']
                        nontrade = row['Average Non-Trade Share']
                        
                        if trade == 1 or nontrade == 1 :
                            # If any share is 100%, don't change
                            return trade,nontrade
                        else:
                            trade = min(trade + 0.05, 1)
                            nontrade = min(nontrade - 0.05, 1 - trade)
                        
                        return trade,nontrade
                    filtered_df['Adjusted Trade Share'], filtered_df['Adjusted Non-Trade Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Trade Share'] * filtered_df['Trade EBITDA'] +
                        filtered_df['Adjusted Non-Trade Share'] * filtered_df['Non-Trade EBITDA']
                    )

                    # Calculate differences
                    filtered_df['T-NT Difference'] = filtered_df['Trade EBITDA'] - filtered_df['Non-Trade EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Trade EBITDA'],
                                             mode='lines+markers', name='Trade EBIDTA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Non-Trade EBITDA'],
                                             mode='lines+markers', name='Non-Trade EBIDTA', line=dict(color='blue')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='crimson', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBIDTA',
                                             line=dict(color='brown', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(T-NT: {g_r:.0f})<br>(I-O: {g_y:.0f}))" 
                                for month, g_r, g_y in 
                                zip(filtered_df['Month'], 
                                    filtered_df['T-NT Difference'],  
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}(Type:-{product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Trade','Non-Trade']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [45,45,45] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                    
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Trade Share', 'Average Non-Trade Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Trade Share': 'green', 'Average Non-Trade Share': 'blue'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Trade', 'Non-Trade']]
                    for _, row in filtered_df[['Month', 'Trade', 'Non-Trade','Average Trade Share', 'Average Non-Trade Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Trade']:.0f} ({row['Average Trade Share']:.2%})",
                            f"{row['Non-Trade']:.0f} ({row['Average Non-Trade Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


if selected == "Home":
    st.title("📊 Advanced Segment Mix Analysis")
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
    st.title("📈 Segment Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                subset_df = df[(df['Region'] == region) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
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
        trade_share = st.sidebar.slider("Adjust Trade Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Trade NSR', 'Non-Trade NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Trade Contribution', 'Non-Trade Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Trade'] * filtered_df[cols[0]] +
                                        filtered_df['Non-Trade'] * filtered_df[cols[1]]) / (
                                            filtered_df['Trade'] + filtered_df['Non-Trade'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - trade_share/100) * filtered_df[cols[1]] +
                                          (trade_share/100) * filtered_df[cols[0]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            for col in cols:
                fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[col],
                                         mode='lines+markers', name=col))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({trade_share}% Trade)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(T-NT: {diff:.2f})<br>(I-O: {i_diff:.2f})" for month, diff, i_diff in 
                        zip(filtered_df['Month'], filtered_df['Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (T-NT: Trade - Non-Trade, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
            
            # Display descriptive statistics
            st.subheader("Descriptive Statistics")
            desc_stats = filtered_df[cols + [overall_col, imaginary_col]].describe()
            st.dataframe(desc_stats.style.format("{:.2f}"), use_container_width=True)
            
            # Display share of Normal and Premium Products
            st.subheader("Share of Trade and Non-Trade Channel")
            total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
            trade_share = (filtered_df['Trade'] / total_quantity * 100).round(2)
            nontrade_share = (filtered_df['Non-Trade'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Trade Share (%)': trade_share,
                'Non-Trade Share (%)': nontrade_share
            })
                  
            fig_pie = px.pie(share_df, values=[trade_share.mean(), nontrade_share.mean()], 
                                     names=['Trade', 'Non-Trade'], title='Average Share Distribution',color=["T","NT"],color_discrete_map={"T":"green","NT":"blue"},hole=0.5)
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
   
