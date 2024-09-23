import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
import base64
from io import BytesIO

# Set page config
st.set_page_config(page_title="Advanced Data Analysis App", page_icon="📊", layout="wide")

# Function to load Lottie animations
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

# Custom CSS
st.markdown("""
<style>
    .main {
        background-color: #f0f2f6;
    }
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .upload-section, .analysis-section, .edit-section {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 10px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-top: 20px;
    }
    .stButton>button {
        width: 100%;
    }
    .stProgress > div > div > div > div {
        background-color: #4CAF50;
    }
    .stSlider > div > div > div > div {
        background-color: #2196F3;
    }
</style>
""", unsafe_allow_html=True)

# Sidebar navigation
with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Data Analysis", "Insights", "Export"],
        icons=["house", "graph-up", "lightbulb", "file-earmark-arrow-down"],
        menu_icon="cast",
        default_index=0,
    )

# Session state
if 'df' not in st.session_state:
    st.session_state.df = None

if selected == "Home":
    st.title("📊 Advanced Product Analysis Dashboard")
    st.markdown("Welcome to our state-of-the-art data analysis tool. Upload your Excel file to get started with interactive visualizations and in-depth insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    with col2:
        if lottie_upload:
            st_lottie(lottie_upload, height=150, key="upload")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/4503/4503700.png", width=150)
    st.markdown("</div>", unsafe_allow_html=True)
    
    if uploaded_file is not None:
        st.session_state.df = pd.read_excel(uploaded_file)
        st.success("File uploaded successfully! Navigate to the Data Analysis section to explore your data.")

elif selected == "Data Analysis":
    if st.session_state.df is not None:
        st.title("Data Analysis")
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        # Display Lottie animation or static image
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        
        # Create sidebar for user inputs
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=['All'] + list(st.session_state.df['Region'].unique()))
        brand = st.sidebar.selectbox("Select Brand", options=['All'] + list(st.session_state.df['Brand'].unique()))
        product_type = st.sidebar.selectbox("Select Type", options=['All'] + list(st.session_state.df['Type'].unique()))
        region_subset = st.sidebar.selectbox("Select Region Subset", options=['All'] + list(st.session_state.df['Region subsets'].unique()))
        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "NSR Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type))
        
        # Update session state
        st.session_state.analysis_type = analysis_type

        premium_share = st.sidebar.slider("Adjust Premium Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = st.session_state.df.copy()
        if region != 'All':
            filtered_df = filtered_df[filtered_df['Region'] == region]
        if brand != 'All':
            filtered_df = filtered_df[filtered_df['Brand'] == brand]
        if product_type != 'All':
            filtered_df = filtered_df[filtered_df['Type'] == product_type]
        if region_subset != 'All':
            filtered_df = filtered_df[filtered_df['Region subsets'] == region_subset]
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Normal NSR', 'Premium NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Normal Contribution', 'Premium Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Normal EBITDA', 'Premium EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Normal Quantity'] * filtered_df[cols[0]] +
                                        filtered_df['Premium Quantity'] * filtered_df[cols[1]]) / (
                                            filtered_df['Normal Quantity'] + filtered_df['Premium Quantity'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - premium_share/100) * filtered_df[cols[0]] +
                                          (premium_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[1]] - filtered_df[cols[0]]
            
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
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({premium_share}% Premium)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(P-N: {diff:.2f})<br>(I-O: {i_diff:.2f})" for month, diff, i_diff in 
                        zip(filtered_df['Month'], filtered_df['Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (P-N: Premium - Normal, I-O: Imaginary - Overall)',
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
            st.subheader("Share of Normal and Premium Products")
            total_quantity = filtered_df['Normal Quantity'] + filtered_df['Premium Quantity']
            normal_share = (filtered_df['Normal Quantity'] / total_quantity * 100).round(2)
            premium_share = (filtered_df['Premium Quantity'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Normal Share (%)': normal_share,
                'Premium Share (%)': premium_share
            })
            
            st.dataframe(share_df.set_index('Month'), use_container_width=True)
            
            # Add a stacked bar chart for product share
            fig_share = px.bar(share_df, x='Month', y=['Normal Share (%)', 'Premium Share (%)'], 
                                title='Monthly Share of Normal and Premium Products',
                                labels={'value': 'Share (%)', 'variable': 'Product Type'},
                                color_discrete_map={'Normal Share (%)': '#1f77b4', 'Premium Share (%)': '#ff7f0e'})
            fig_share.update_layout(barmode='stack')
            st.plotly_chart(fig_share, use_container_width=True)
            
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    else:
        st.info("Please upload an Excel file in the Home section to begin the analysis.")

elif selected == "Insights":
    if st.session_state.df is not None:
        st.title("Data Insights")
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        # Overall trends
        st.subheader("Overall Trends")
        total_nsr = st.session_state.df['Normal NSR'] + st.session_state.df['Premium NSR']
        total_contribution = st.session_state.df['Normal Contribution'] + st.session_state.df['Premium Contribution']
        total_ebitda = st.session_state.df['Normal EBITDA'] + st.session_state.df['Premium EBITDA']
        
        col1, col2, col3 = st.columns(3)
        col1.metric("Total NSR", f"${total_nsr.sum():,.2f}", f"{total_nsr.pct_change().mean()*100:.2f}%")
        col2.metric("Total Contribution", f"${total_contribution.sum():,.2f}", f"{total_contribution.pct_change().mean()*100:.2f}%")
        col3.metric("Total EBITDA", f"${total_ebitda.sum():,.2f}", f"{total_ebitda.pct_change().mean()*100:.2f}%")
        
        # Top performing products
        st.subheader("Top Performing Products")
        top_products = st.session_state.df.groupby('Brand')['Premium NSR'].sum().sort_values(ascending=False).head(5)
        fig_top = px.bar(top_products, x=top_products.index, y='Premium NSR', title='Top 5 Products by Premium NSR')
        st.plotly_chart(fig_top, use_container_width=True)
        
        # Regional performance
        st.subheader("Regional Performance")
        regional_perf = st.session_state.df.groupby('Region')[['Normal NSR', 'Premium NSR']].sum()
        fig_region = px.bar(regional_perf, x=regional_perf.index, y=['Normal NSR', 'Premium NSR'], 
                            title='NSR by Region', barmode='group')
        st.plotly_chart(fig_region, use_container_width=True)
        
        # Correlation analysis
        st.subheader("Correlation Analysis")
        corr_matrix = st.session_state.df[['Normal NSR', 'Premium NSR', 'Normal Contribution', 'Premium Contribution', 'Normal EBITDA', 'Premium EBITDA']].corr()
        fig_corr = px.imshow(corr_matrix, text_auto=True, aspect="auto", title='Correlation Matrix')
        st.plotly_chart(fig_corr, use_container_width=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Please upload an Excel file in the Home section to view insights.")


elif selected == "Export":
    if st.session_state.df is not None:
        st.title("Export Data and Visualizations")
        st.markdown("<div class='export-section'>", unsafe_allow_html=True)
        
        # Export filtered data
        st.subheader("Export Filtered Data")
        export_format = st.selectbox("Select export format", ["CSV", "Excel"])
        
        if st.button("Export Filtered Data"):
            filtered_df = st.session_state.df  # You may want to use the filtered dataframe from the analysis section
            if export_format == "CSV":
                csv = filtered_df.to_csv(index=False)
                b64 = base64.b64encode(csv.encode()).decode()
                href = f'<a href="data:file/csv;base64,{b64}" download="filtered_data.csv">Download CSV File</a>'
                st.markdown(href, unsafe_allow_html=True)
            else:
                towrite = BytesIO()
                filtered_df.to_excel(towrite, index=False, engine='openpyxl')
                towrite.seek(0)
                b64 = base64.b64encode(towrite.read()).decode()
                href = f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="filtered_data.xlsx">Download Excel File</a>'
                st.markdown(href, unsafe_allow_html=True)
        
        # Export visualizations
        st.subheader("Export Visualizations")
        if st.button("Export All Visualizations"):
            # Create a zip file containing all visualizations
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                # Add main analysis plot
                fig = create_main_analysis_plot(filtered_df)  # You need to implement this function
                img_bytes = fig.to_image(format="png")
                zip_file.writestr("main_analysis_plot.png", img_bytes)
                
                # Add product share plot
                fig_share = create_product_share_plot(filtered_df)  # You need to implement this function
                img_bytes = fig_share.to_image(format="png")
                zip_file.writestr("product_share_plot.png", img_bytes)
                
                # Add top products plot
                fig_top = create_top_products_plot(st.session_state.df)  # You need to implement this function
                img_bytes = fig_top.to_image(format="png")
                zip_file.writestr("top_products_plot.png", img_bytes)
                
                # Add regional performance plot
                fig_region = create_regional_performance_plot(st.session_state.df)  # You need to implement this function
                img_bytes = fig_region.to_image(format="png")
                zip_file.writestr("regional_performance_plot.png", img_bytes)
                
                # Add correlation matrix plot
                fig_corr = create_correlation_matrix_plot(st.session_state.df)  # You need to implement this function
                img_bytes = fig_corr.to_image(format="png")
                zip_file.writestr("correlation_matrix_plot.png", img_bytes)
            
            zip_buffer.seek(0)
            b64 = base64.b64encode(zip_buffer.getvalue()).decode()
            href = f'<a href="data:application/zip;base64,{b64}" download="visualizations.zip">Download All Visualizations</a>'
            st.markdown(href, unsafe_allow_html=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
    else:
        st.info("Please upload an Excel file in the Home section to export data and visualizations.")

# Add a footer
st.markdown("---")
st.markdown("Created with ❤️ by Prasoon Bajpai")

# Helper functions for creating plots (you need to implement these)
def create_main_analysis_plot(df):
    # Implement the main analysis plot creation here
    pass

def create_product_share_plot(df):
    # Implement the product share plot creation here
    pass

def create_top_products_plot(df):
    # Implement the top products plot creation here
    pass

def create_regional_performance_plot(df):
    # Implement the regional performance plot creation here
    pass

def create_correlation_matrix_plot(df):
    # Implement the correlation matrix plot creation here
    pass
