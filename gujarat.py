import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from io import BytesIO
import base64

# New function to create PDF report
def create_pdf_report(region, df):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas, doc):
        canvas.saveState()
        canvas.setFont('Helvetica', 10)
        page_number_text = f"Page {doc.page}"
        canvas.drawString(width - 100, 30, page_number_text)
        canvas.restoreState()

    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png")
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    # Title
    c.setFont("Helvetica-Bold", 24)
    c.drawString(50, height - 50, f"GYR Analysis Report for {region}")

    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_count = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    # EBITDA Analysis
                    cols = ['Green EBITDA', 'Yellow EBITDA', 'Red EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    filtered_df[overall_col] = (filtered_df['Green'] * filtered_df[cols[0]] +
                                                filtered_df['Yellow'] * filtered_df[cols[1]] + 
                                                filtered_df['Red'] * filtered_df[cols[2]]) / (
                                                filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red'])

                    # Calculate imaginary overall based on adjusted shares
                    filtered_df['Current Green Share'] = filtered_df['Green'] / (filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red'])
                    filtered_df['Current Yellow Share'] = filtered_df['Yellow'] / (filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red'])
                    filtered_df['Current Red Share'] = filtered_df['Red'] / (filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red'])

                    filtered_df['Adjusted Green Share'] = filtered_df['Current Green Share'].apply(lambda x: min(x + 0.05, 1) if x > 0 else 0)
                    filtered_df['Adjusted Yellow Share'] = filtered_df['Current Yellow Share'].apply(lambda x: min(x + 0.025, 1 - filtered_df['Adjusted Green Share']) if x > 0 else 0)
                    filtered_df['Adjusted Red Share'] = 1 - filtered_df['Adjusted Green Share'] - filtered_df['Adjusted Yellow Share']

                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Green Share'] * filtered_df['Green EBITDA'] +
                        filtered_df['Adjusted Yellow Share'] * filtered_df['Yellow EBITDA'] +
                        filtered_df['Adjusted Red Share'] * filtered_df['Red EBITDA']
                    )

                    # Create the plot
                    fig = go.Figure()

                    for col in cols:
                        fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[col],
                                                 mode='lines+markers', name=col))

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(dash='dash')))

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBITDA',
                                             line=dict(color='brown', dash='dot')))

                    fig.update_layout(
                        title=f"EBITDA Analysis: {brand} - {product_type} - {region_subset}",
                        xaxis_title='Month',
                        yaxis_title='EBITDA',
                        legend_title='Metrics'
                    )

                    # Add new page if needed
                    if page_count > 1:
                        c.showPage()
                    
                    # Draw the graph
                    draw_graph(fig, 50, height - 500, 500, 400)

                    # Add descriptive statistics
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(50, height - 520, "Descriptive Statistics")
                    desc_stats = filtered_df[cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    for i, row in desc_stats.iterrows():
                        c.setFont("Helvetica", 10)
                        y_position = height - 540 - (i * 15)
                        c.drawString(50, y_position, f"{row['index']}: {', '.join([f'{col}: {row[col]:.2f}' for col in desc_stats.columns if col != 'index'])}")

                    # Add share of Green, Yellow, and Red Products
                    c.setFont("Helvetica-Bold", 14)
                    c.drawString(50, height - 680, "Share of Green, Yellow, and Red Products")
                    
                    # Create pie chart
                    share_fig = px.pie(values=[filtered_df['Current Green Share'].mean(), 
                                               filtered_df['Current Yellow Share'].mean(), 
                                               filtered_df['Current Red Share'].mean()], 
                                       names=['Green', 'Yellow', 'Red'], 
                                       title='Average Share Distribution')
                    
                    draw_graph(share_fig, 50, height - 900, 300, 200)

                    # Add share table
                    c.setFont("Helvetica", 10)
                    for i, (_, row) in enumerate(filtered_df[['Month', 'Current Green Share', 'Current Yellow Share', 'Current Red Share']].iterrows()):
                        y_position = height - 920 - (i * 15)
                        c.drawString(400, y_position, f"{row['Month']}: G: {row['Current Green Share']:.2%}, Y: {row['Current Yellow Share']:.2%}, R: {row['Current Red Share']:.2%}")

                    add_page_number(c, c._pageNumber)
                    page_count += 1

    c.save()
    buffer.seek(0)
    return buffer
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
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique())

        # Add download button for combined report
        if st.sidebar.button(f"Download Combined Report for {region}"):
            pdf_buffer = create_pdf_report(region, df)
            pdf_bytes = pdf_buffer.getvalue()
            b64 = base64.b64encode(pdf_bytes).decode()
            href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}.pdf">Download PDF Report</a>'
            st.sidebar.markdown(href, unsafe_allow_html=True)

        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique())
        brand = st.sidebar.selectbox("Select Brand", options=df['Brand'].unique())
        product_type = st.sidebar.selectbox("Select Type", options=df['Type'].unique())
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df['Region subsets'].unique())
        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type))
        
        # Update session state
        st.session_state.analysis_type = analysis_type

        green_share = st.sidebar.slider("Adjust Green Share (%)", 0, 99, 50)
        yellow_share = st.sidebar.slider("Adjust Yellow Share (%)", 0, 100-green_share,0)
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
   
