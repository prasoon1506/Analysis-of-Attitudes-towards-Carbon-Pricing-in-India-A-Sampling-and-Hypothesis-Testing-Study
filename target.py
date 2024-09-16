import streamlit as st
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import math
from scipy import stats
import matplotlib.pyplot as plt
import seaborn as sns
import xgboost as xgb
from io import BytesIO
import base64
import time
import requests
from streamlit_lottie import st_lottie
from concurrent.futures import ThreadPoolExecutor

# Cache the data loading
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    regions = df['Zone'].unique().tolist()
    brands = df['Brand'].unique().tolist()
    return df, regions, brands

# Cache the model training
@st.cache_resource
def train_model(X_train, y_train):
    model = xgb.XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42)
    model.fit(X_train, y_train)
    return model

def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

def predict_and_visualize(df, region, brand):
    try:
        region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)].copy()
        
        if len(region_data) > 0:
            months = ['Apr', 'May', 'June', 'July', 'Aug']
            for month in months:
                region_data[f'Achievement({month})'] = region_data[f'Monthly Achievement({month})'] / region_data[f'Month Tgt ({month})']
            
            X = region_data[[f'Month Tgt ({month})' for month in months]]
            y = region_data[[f'Achievement({month})' for month in months]]
            
            X_reshaped = X.values.reshape(-1, 1)
            y_reshaped = y.values.ravel()
            
            X_train, X_val, y_train, y_val = train_test_split(X_reshaped, y_reshaped, test_size=0.2, random_state=42)
            
            model = train_model(X_train, y_train)
            
            val_predictions = model.predict(X_val)
            rmse = math.sqrt(mean_squared_error(y_val, val_predictions))
            
            sept_target = region_data['Month Tgt (Sep)'].iloc[-1]
            sept_prediction = model.predict([[sept_target]])[0]
            
            n = len(X_train)
            degrees_of_freedom = n - 2
            t_value = stats.t.ppf(0.975, degrees_of_freedom)
            
            residuals = y_train - model.predict(X_train)
            std_error = np.sqrt(np.sum(residuals**2) / degrees_of_freedom)
            
            margin_of_error = t_value * std_error * np.sqrt(1 + 1/n + (sept_target - np.mean(X_train))**2 / np.sum((X_train - np.mean(X_train))**2))
            
            lower_bound = max(0, sept_prediction - margin_of_error)
            upper_bound = sept_prediction + margin_of_error
            
            sept_achievement = sept_prediction * sept_target
            lower_achievement = lower_bound * sept_target
            upper_achievement = upper_bound * sept_target
            
            fig = create_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse)
            
            return fig, sept_achievement, lower_achievement, upper_achievement, rmse
        else:
            return None, None, None, None, None
    except Exception as e:
        st.error(f"Error in predict_and_visualize: {str(e)}")
        raise

def create_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse):
    fig = plt.figure(figsize=(16, 18))
    gs = fig.add_gridspec(4, 1, height_ratios=[0.5, 0.5, 3, 1])
    
    ax_region = fig.add_subplot(gs[0])
    ax_region.axis('off')
    ax_region.text(0.5, 0.5, region, fontsize=24, fontweight='bold', ha='center', va='center')
    
    ax_table = fig.add_subplot(gs[1])
    ax_table.axis('off')
    table_data = [
        ['Brand', 'Month Target (Sep)', 'Monthly Achievement (Aug)', 'Predicted Achievement', 'CI', 'RMSE'],
        [brand, f"{sept_target:.2f}", f"{region_data['Monthly Achievement(Aug)'].iloc[-1]:.2f}", 
         f"{sept_achievement:.2f}", f"({lower_achievement:.2f}, {upper_achievement:.2f})", f"{rmse:.4f}"]
    ]
    table = ax_table.table(cellText=table_data, colLabels=None, cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.2)
    for (row, col), cell in table.get_celld().items():
        if row == 0:
            cell.set_text_props(fontweight='bold')
    
    ax1 = fig.add_subplot(gs[2])
    
    actual_achievements = [region_data[f'Monthly Achievement({month})'].iloc[-1] for month in months]
    actual_targets = [region_data[f'Month Tgt ({month})'].iloc[-1] for month in months]
    all_months = months + ['Sep']
    all_achievements = actual_achievements + [sept_achievement]
    all_targets = actual_targets + [sept_target]
    
    x = np.arange(len(all_months))
    width = 0.35
    
    rects1 = ax1.bar(x - width/2, all_targets, width, label='Target', color='pink', alpha=0.8)
    rects2 = ax1.bar(x + width/2, all_achievements, width, label='Achievement', color='yellow', alpha=0.8)
    
    ax1.bar(x[-1] + width/2, sept_achievement, width, color='red', alpha=0.8)
    
    ax1.set_ylabel('Target and Achievement', fontsize=12, fontweight='bold')
    ax1.set_title(f"Monthly Targets and Achievements for FY 2025", fontsize=18, fontweight='bold')
    ax1.set_xticks(x)
    ax1.set_xticklabels(all_months)
    ax1.legend()
    
    def autolabel(rects):
        for rect in rects:
            height = rect.get_height()
            ax1.annotate(f'{height:.0f}',
                        xy=(rect.get_x() + rect.get_width() / 2, height),
                        xytext=(0, 3),
                        textcoords="offset points",
                        ha='center', va='bottom', fontsize=8)
    
    autolabel(rects1)
    autolabel(rects2)
    
    for i, (target, achievement) in enumerate(zip(all_targets, all_achievements)):
        percentage = (achievement / target) * 100
        color = 'green' if percentage >= 100 else 'red'
        ax1.text(i, max(target, achievement), f'{percentage:.1f}%', 
                 ha='center', va='bottom', fontsize=10, color=color, fontweight='bold')
    
    ax1.errorbar(x[-1] + width/2, sept_achievement, 
                 yerr=[[sept_achievement - lower_achievement], [upper_achievement - sept_achievement]],
                 fmt='o', color='darkred', capsize=5, capthick=2, elinewidth=2)
    
    ax2 = fig.add_subplot(gs[3])
    percent_achievements = [((ach / tgt) * 100) for ach, tgt in zip(all_achievements, all_targets)]
    ax2.plot(x, percent_achievements, marker='o', linestyle='-', color='purple')
    ax2.axhline(y=100, color='r', linestyle='--', alpha=0.7)
    ax2.set_xlabel('Month', fontsize=12, fontweight='bold')
    ax2.set_ylabel('% Achievement', fontsize=12, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(all_months)
    
    for i, pct in enumerate(percent_achievements):
        ax2.annotate(f'{pct:.1f}%', (i, pct), xytext=(0, 5), textcoords='offset points', 
                     ha='center', va='bottom', fontsize=8)
    
    plt.tight_layout()
    return fig
def generate_combined_report(df, regions, brands):
    table_data = [['Region', 'Brand', 'Month Target\n(Sep)', 'Monthly Achievement\n(Aug)', 'Predicted\nAchievement', 'CI', 'RMSE']]
    
    with ThreadPoolExecutor() as executor:
        futures = []
        for region in regions:
            for brand in brands:
                futures.append(executor.submit(predict_and_visualize, df, region, brand))
        
        for future, (region, brand) in zip(futures, [(r, b) for r in regions for b in brands]):
            _, sept_achievement, lower_achievement, upper_achievement, rmse = future.result()
            if sept_achievement is not None:
                region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)]
                if not region_data.empty:
                    sept_target = region_data['Month Tgt (Sep)'].iloc[-1]
                    aug_achievement = region_data['Monthly Achievement(Aug)'].iloc[-1]
                    
                    table_data.append([
                        region, brand, f"{sept_target:.2f}", f"{aug_achievement:.2f}",
                        f"{sept_achievement:.2f}", f"({lower_achievement:.2f},\n{upper_achievement:.2f})", f"{rmse:.4f}"
                    ])
                else:
                    st.warning(f"No data available for {region} and {brand}")
    
    if len(table_data) > 1:
        # Determine the number of rows in the table
        num_rows = len(table_data)
        
        # Calculate the figure height based on the number of rows
        fig_height = max(6, 1 + 0.5 * num_rows)  # Minimum height of 6, scales with number of rows
        
        fig, ax = plt.subplots(figsize=(12, fig_height))
        ax.axis('off')
        
        # Add title to the figure, not the axis
        fig.suptitle("", fontsize=16, fontweight='bold', y=0.95)
        
        # Create the table
        table = ax.table(cellText=table_data[1:], colLabels=table_data[0], cellLoc='center', loc='center')
        
        # Set font size and style
        table.auto_set_font_size(False)
        table.set_fontsize(8)
        
        # Adjust column widths
        col_widths = [0.15, 0.15, 0.15, 0.15, 0.15, 0.15, 0.1]
        for i, width in enumerate(col_widths):
            table.auto_set_column_width(i)
            
        # Style the header
        for (row, col), cell in table.get_celld().items():
            if row == 0:
                cell.set_text_props(fontweight='bold', wrap=True)
                cell.set_height(0.1)
            else:
                cell.set_height(0.05)
        
        # Adjust the layout
        table.scale(1, 1.5)
        plt.subplots_adjust(top=0.9, bottom=0.02, left=0.05, right=0.95)
        
        pdf_buffer = BytesIO()
        fig.savefig(pdf_buffer, format='pdf', bbox_inches='tight')
        plt.close(fig)
        
        pdf_buffer.seek(0)
        return base64.b64encode(pdf_buffer.getvalue()).decode()
    else:
        st.warning("No data available for any region and brand combination.")
        return None

# The main function remains the same as in the previous response


def main():
    st.set_page_config(page_title="Sales Prediction App", page_icon="📊", layout="wide")
    
    # Load Lottie animation
    lottie_url = "https://assets5.lottiefiles.com/packages/lf20_V9t630.json"
    lottie_json = load_lottie_url(lottie_url)
    
    # Sidebar
    with st.sidebar:
        st_lottie(lottie_json, height=200)
        st.title("Navigation")
        page = st.radio("Go to", ["Home", "Predictions", "About"])
    
    if page == "Home":
        st.title("📊 Welcome to the Sales Prediction App")
        st.write("This app helps you predict and visualize sales achievements for different regions and brands.")
        st.write("Use the sidebar to navigate between pages and upload your data to get started!")
        
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
        if uploaded_file is not None:
            with st.spinner("Loading data..."):
                df, regions, brands = load_data(uploaded_file)
            st.session_state['df'] = df
            st.session_state['regions'] = regions
            st.session_state['brands'] = brands
            st.success("File uploaded and processed successfully!")
    
    elif page == "Predictions":
        st.title("🔮 Sales Predictions")
        if 'df' not in st.session_state:
            st.warning("Please upload a file on the Home page first.")
        else:
            df = st.session_state['df']
            regions = st.session_state['regions']
            brands = st.session_state['brands']
            
            col1, col2 = st.columns(2)
            with col1:
                region = st.selectbox("Select Region", regions)
            with col2:
                brand = st.selectbox("Select Brand", brands)
            
            if st.button("Run Prediction"):
                with st.spinner("Running prediction..."):
                    fig, sept_achievement, lower_achievement, upper_achievement, rmse = predict_and_visualize(df, region, brand)
                if fig:
                    st.pyplot(fig)
                    
                    # Individual report download
                    buf = BytesIO()
                    fig.savefig(buf, format="pdf")
                    buf.seek(0)
                    b64 = base64.b64encode(buf.getvalue()).decode()
                    st.download_button(
                        label="Download Individual PDF Report",
                        data=buf,
                        file_name=f"prediction_report_{region}_{brand}.pdf",
                        mime="application/pdf"
                    )
                else:
                    st.error(f"No data available for {region} and {brand}")
            

            if st.button("Generate Combined Report"):
              with st.spinner("Generating combined report..."):
                combined_report_data = generate_combined_report(df, regions, brands)
              if combined_report_data:
                st.download_button(
                    label="Download Combined PDF Report",
                    data=base64.b64decode(combined_report_data),
                    file_name="combined_prediction_report.pdf",
                    mime="application/pdf"
                )
              else:
                st.error("Unable to generate combined report due to lack of data.")
    
    elif page == "About":
        st.title("ℹ️ About the Sales Prediction App")
        st.write("""
        This app is designed to help sales teams predict and visualize their performance across different regions and brands.
        
        Key features:
        - Data upload and processing
        - Individual predictions for each region and brand
        - Combined report generation
        - Interactive visualizations
        
        For any questions or support, please contact our team at support@salespredictionapp.com
        """)

if __name__ == "__main__":
    main()
