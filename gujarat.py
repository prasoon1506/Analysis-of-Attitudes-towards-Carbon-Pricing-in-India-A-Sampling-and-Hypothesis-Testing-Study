import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu

# Set page config
st.set_page_config(page_title="Advanced Data Analysis", page_icon="📊", layout="wide")

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
@import url('https://fonts.googleapis.com/css2?family=Roboto:wght@100;300;400;700&display=swap');

html, body, [class*="css"] {
    font-family: 'Roboto', sans-serif;
}

.main {
    background-color: #f0f2f6;
}

.stApp {
    max-width: 1200px;
    margin: 0 auto;
}

.upload-section, .analysis-section, .edit-section {
    background-color: #ffffff;
    padding: 30px;
    border-radius: 15px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    margin-top: 30px;
    transition: all 0.3s ease;
}

.upload-section:hover, .analysis-section:hover, .edit-section:hover {
    box-shadow: 0 10px 20px rgba(0, 0, 0, 0.15);
}

.stButton>button {
    width: 100%;
    border-radius: 5px;
    font-weight: 500;
    transition: all 0.3s ease;
}

.stButton>button:hover {
    background-color: #4CAF50;
    color: white;
}

h1, h2, h3 {
    color: #2C3E50;
}

.stPlotlyChart {
    background-color: white;
    border-radius: 10px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    padding: 20px;
    margin-top: 20px;
}

.stDataFrame {
    border-radius: 10px;
    box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
}

</style>
""", unsafe_allow_html=True)





# Sidebar navigation
with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )

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
        brand = st.sidebar.selectbox("Select Brand", options=df['Brand'].unique())
        product_type = st.sidebar.selectbox("Select Type", options=df['Type'].unique())
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df['Region subsets'].unique())
        
        # Move sliders to sidebar and update session state
        st.sidebar.header("Adjust Shares")
        green_share = st.sidebar.slider("Green Share (%)", 0, 99, step=1, value=st.session_state.green_share, key="green_slider")
        yellow_share = st.sidebar.slider("Yellow Share (%)", 0, 100-green_share, step=1, value=st.session_state.yellow_share, key="yellow_slider")
        
        # Update session state
        st.session_state.green_share = green_share
        st.session_state.yellow_share = yellow_share
        
        
        # Display red share
        red_share = 100 - green_share - yellow_share
        st.sidebar.text(f"Red Share: {red_share}%")
        
        # Analysis type selection using tabs
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        analysis_type = st.tabs(analysis_options)
        
        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            for idx, analysis in enumerate(analysis_options):
                with analysis_type[idx]:
                    if analysis == 'NSR Analysis':
                        cols = ['Green NSR', 'Yellow NSR', 'Red NSR']
                        overall_col = 'Overall NSR'
                    elif analysis == 'Contribution Analysis':
                        cols = ['Green Contribution', 'Yellow Contribution', 'Red Contribution']
                        overall_col = 'Overall Contribution'
                    elif analysis == 'EBITDA Analysis':
                        cols = ['Green EBITDA', 'Yellow EBITDA', 'Red EBITDA']
                        overall_col = 'Overall EBITDA'
                    
                    # Calculate weighted average based on actual quantities
                    filtered_df[overall_col] = (filtered_df['Green'] * filtered_df[cols[0]] +
                                                filtered_df['Yellow'] * filtered_df[cols[1]] + 
                                                filtered_df['Red'] * filtered_df[cols[2]]) / (
                                                    filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red'])
                    
                    # Calculate imaginary overall based on slider
                    imaginary_col = f'Imaginary {overall_col}'
                    filtered_df[imaginary_col] = ((red_share / 100) * filtered_df[cols[2]] +
                                                  (green_share / 100) * filtered_df[cols[0]] + 
                                                  (yellow_share / 100) * filtered_df[cols[1]])
                    
                    # Calculate differences
                    filtered_df['G-Y Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
                    filtered_df['G-R Difference'] = filtered_df[cols[0]] - filtered_df[cols[2]]
                    filtered_df['Y-R Difference'] = filtered_df[cols[1]] - filtered_df[cols[2]]
                    filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    
                    for col in cols:
                        fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[col],
                                                 mode='lines+markers', name=col))
                    
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(dash='dash')))
                    
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                             mode='lines+markers', 
                                             name=f'Imaginary {overall_col} ({green_share}% Green, {yellow_share}% Yellow, {red_share}% Red)',
                                             line=dict(color='brown', dash='dot')))
                    
                    # Customize x-axis labels to include the differences
                    x_labels = [f"{month}<br>(G-Y: {diff:.2f})<br>(G-R: {i_diff:.2f})<br>(Y-R: {j_diff:.2f})<br>(I-O: {k_diff:.2f})" 
                                for month, diff, i_diff, j_diff, k_diff in 
                                zip(filtered_df['Month'], filtered_df['G-Y Difference'], 
                                    filtered_df['G-R Difference'], filtered_df['Y-R Difference'], 
                                    filtered_df['Imaginary vs Overall Difference'])]
                    
                    fig.update_layout(
                        title=analysis,
                        xaxis_title='Month',
                        yaxis_title='Value',
                        legend_title='Metrics',
                        hovermode="x unified",
                        xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
                    )
                    
                    st.plotly_chart(fig, use_container_width=True)
                    
                    # Display descriptive statistics
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
                                     names=['Green', 'Yellow', 'Red'], title='Average Share Distribution')
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
    
    For more information or support, please contact our team.
    """)
