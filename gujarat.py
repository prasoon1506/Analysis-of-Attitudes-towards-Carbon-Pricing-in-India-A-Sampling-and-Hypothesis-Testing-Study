import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import requests
from streamlit_lottie import st_lottie
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.preprocessing import LabelEncoder

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
.upload-section, .analysis-section, .edit-section, .prediction-section {
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
st.title("📊 Normal Vs. Premium Product Analysis")
st.markdown("Upload your Excel file and analyze it with interactive visualizations.")
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
    
    # Create sidebar for user inputs with cascading filters
    st.sidebar.header("Filter Options")
    region = st.sidebar.selectbox("Select Region", options=['All'] + list(df['Region'].unique()))
    
    # Filter brands based on selected region
    if region != 'All':
        brand_options = ['All'] + list(df[df['Region'] == region]['Brand'].unique())
    else:
        brand_options = ['All'] + list(df['Brand'].unique())
    brand = st.sidebar.selectbox("Select Brand", options=brand_options)
    
    # Filter types based on selected region and brand
    if region != 'All' and brand != 'All':
        type_options = ['All'] + list(df[(df['Region'] == region) & (df['Brand'] == brand)]['Type'].unique())
    elif region != 'All':
        type_options = ['All'] + list(df[df['Region'] == region]['Type'].unique())
    elif brand != 'All':
        type_options = ['All'] + list(df[df['Brand'] == brand]['Type'].unique())
    else:
        type_options = ['All'] + list(df['Type'].unique())
    product_type = st.sidebar.selectbox("Select Type", options=type_options)
    
    # Filter region subsets based on selected region
    if region != 'All':
        region_subset_options = ['All'] + list(df[df['Region'] == region]['Region subsets'].unique())
    else:
        region_subset_options = ['All'] + list(df['Region subsets'].unique())
    region_subset = st.sidebar.selectbox("Select Region Subset", options=region_subset_options)
    
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
    filtered_df = df.copy()
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
        
        # ML Prediction Section
        st.markdown("<div class='prediction-section'>", unsafe_allow_html=True)
        st.subheader("Advanced Prediction")
        
        # Prepare data for ML model
        le = LabelEncoder()
        df_ml = df.copy()
        df_ml['Month'] = pd.to_datetime(df_ml['Month'])
        df_ml['Month_num'] = df_ml['Month'].dt.month + df_ml['Month'].dt.year * 12
        df_ml['Region'] = le.fit_transform(df_ml['Region'])
        df_ml['Brand'] = le.fit_transform(df_ml['Brand'])
        df_ml['Type'] = le.fit_transform(df_ml['Type'])
        df_ml['Region subsets'] = le.fit_transform(df_ml['Region subsets'])
        
        features = ['Month_num', 'Region', 'Brand', 'Type', 'Region subsets']
        target_cols = cols + [overall_col]
        
        # Train ML models
        models = {}
        for target in target_cols:
            X = df_ml[features]
            y = df_ml[target]
            X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
            model = RandomForestRegressor(n_estimators=100, random_state=42)
            model.fit(X_train, y_train)
            models[target] = model
        
        # Predict future values
        last_month = df_ml['Month_num'].max()
        future_months = pd.DataFrame({'Month_num': range(last_month + 1, last_month + 7)})
        future_months['Region'] = df_ml['Region'].mode().iloc[0]
        future_months['Brand'] = df_ml['Brand'].mode().iloc[0]
        future_months['Type'] = df_ml['Type'].mode().iloc[0]
        future_months['Region subsets'] = df_ml['Region subsets'].mode().iloc[0]
        
        predictions = {}
        for target, model in models.items():
            predictions[target] = model.predict(future_months)
        
        # Display predictions
        pred_df = pd.DataFrame(predictions)
        pred_df['Month'] = pd.date_range(start=df['Month'].max() + pd.DateOffset(months=1), periods=6, freq='M')
        pred_df = pred_df.set_index('Month')
        
        st.write("Predicted values for the next 6 months:")
        st.dataframe(pred_df.style.format("{:.2f}"), use_container_width=True)
        
        # Plot predictions
        fig_pred = go.Figure()
        
        for col in target_cols:
            fig_pred.add_trace(go.Scatter(x=pred_df.index, y=pred_df[col],
                                          mode='lines+markers', name=f'Predicted {col}'))
        
        fig_pred.update_layout(
            title="Predicted Values",
            xaxis_title='Month',
            yaxis_title='Value',
            legend_title='Metrics',
            hovermode="x unified"
        )
        
        st.plotly_chart(fig_pred, use_container_width=True)
        
        st.markdown("</div>", unsafe_allow_html=True)
        
    else:
        st.warning("No data available for the selected combination.")
    
    st.markdown("</div>", unsafe_allow_html=True)
    
else:
    st.info("Please upload an Excel file to begin the analysis.")

# Add a footer
st.markdown("---")
st.markdown("Created with ❤️ by Prasoon Bajpai")
