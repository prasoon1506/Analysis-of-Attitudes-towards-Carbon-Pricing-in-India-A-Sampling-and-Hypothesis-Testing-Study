import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import requests
from streamlit_lottie import st_lottie

# Set page config
st.set_page_config(page_title="Data Analysis App", page_icon="📊", layout="wide")

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
</style>
""", unsafe_allow_html=True)

# Title and description
st.title("📊 Interactive Data Analysis App")
st.markdown("Upload your Excel file, edit the data, and analyze it with interactive visualizations.")
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
    # Read the Excel file
        df = pd.read_excel(uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        # Display Lottie animation or static image
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        
        # Create sidebar for user inputs
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique())
        brand = st.sidebar.selectbox("Select Brand", options=df['Brand'].unique())
        product_type = st.sidebar.selectbox("Select Type", options=df['Type'].unique())
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df['Region subsets'].unique())
        
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
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
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
            
        else:
            st.warning("No data available for the selected combination.")
        
        st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown("</div>", unsafe_allow_html=True)

else:
    st.info("Please upload an Excel file to begin the analysis.")

# Add a footer
st.markdown("---")
st.markdown("Created with ❤️ by Prasoon Bajpai")
