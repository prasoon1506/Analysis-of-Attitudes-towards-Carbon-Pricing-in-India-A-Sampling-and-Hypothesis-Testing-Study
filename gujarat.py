import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import requests
from streamlit_lottie import st_lottie
from streamlit.components.v1 import html

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

# Vertical slider component
def vertical_slider(value, min_value, max_value, step, key):
    slider_html = f"""
        <input type="range" min="{min_value}" max="{max_value}" value="{value}" step="{step}" 
               style="width: 300px; writing-mode: bt-lr; -webkit-appearance: slider-vertical; transform: rotate(270deg);">
        <p id="slider-value">{value}%</p>
        <script>
            var slider = document.querySelector('input[type="range"]');
            var output = document.getElementById("slider-value");
            slider.oninput = function() {{
                output.innerHTML = this.value + "%";
                parent.postMessage({{
                    type: "streamlit:setComponentValue",
                    value: this.value
                }}, "*");
            }}
        </script>
    """
    component_value = html(slider_html, height=350, key=key)
    return int(component_value) if component_value else value

# Main app
def main():
    st.title("📊 Interactive Data Analysis App")
    
    # Create tabs
    tab1, tab2, tab3 = st.tabs(["Upload Data", "Analyze Data", "Other Tab"])
    
    with tab1:
        st.header("Upload Your Data")
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
        if uploaded_file is not None:
            st.session_state.df = pd.read_excel(uploaded_file)
            st.success("File uploaded successfully!")
    
    with tab2:
        if 'df' not in st.session_state:
            st.info("Please upload an Excel file in the 'Upload Data' tab to begin the analysis.")
        else:
            analyze_data(st.session_state.df)
    
    with tab3:
        st.header("Other Tab Content")
        st.write("This is another tab where you can add more functionality to your app.")

def analyze_data(df):
    st.header("Data Analysis")
    
    # Create sidebar for user inputs
    st.sidebar.header("Filter Options")
    region = st.sidebar.selectbox("Select Region", options=df['Region'].unique())
    brand = st.sidebar.selectbox("Select Brand", options=df['Brand'].unique())
    product_type = st.sidebar.selectbox("Select Type", options=df['Type'].unique())
    region_subset = st.sidebar.selectbox("Select Region Subset", options=df['Region subsets'].unique())
    
    # Analysis type selection using radio buttons
    st.sidebar.header("Analysis on")
    analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
    analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options)

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
        
        # Create two columns: one for the graph and one for the slider
        col1, col2 = st.columns([4, 1])
        
        with col2:
            premium_share = vertical_slider(50, 0, 100, 1, "premium_share")
        
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
        
        with col1:
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

# Run the app
if __name__ == "__main__":
    main()

# Add a footer
st.markdown("---")
st.markdown("Created with ❤️ by Prasoon Bajpai")
