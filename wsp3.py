import streamlit as st
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import matplotlib.backends.backend_pdf

# Set page config
st.set_page_config(page_title="District Price Trend Analysis", layout="wide")

# Custom CSS
st.markdown("""
<style>
    .stApp {
        max-width: 1200px;
        margin: 0 auto;
    }
    .main .block-container {
        padding-top: 2rem;
    }
    h1 {
        color: #2c3e50;
    }
    .stSelectbox, .stMultiSelect {
        margin-bottom: 1rem;
    }
    .stButton > button {
        width: 100%;
    }
    .stSlider > div > div > div {
        background-color: #3498db;
    }
</style>
""", unsafe_allow_html=True)

# Global variables
if 'df' not in st.session_state:
    st.session_state.df = None
if 'week_names_input' not in st.session_state:
    st.session_state.week_names_input = []
if 'desired_diff_input' not in st.session_state:
    st.session_state.desired_diff_input = {}
if 'file_processed' not in st.session_state:
    st.session_state.file_processed = False
if 'diff_week' not in st.session_state:
    st.session_state.diff_week = 0
if 'district_benchmark_brands' not in st.session_state:
    st.session_state.district_benchmark_brands = {}

def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]
        
        # Renaming columns with a unique suffix
        week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        }, inplace=True)
        
        week_data.replace(0, np.nan, inplace=True)
        
        # Checking if columns already exist in the transformed_df
        existing_columns = set(transformed_df.columns)
        new_columns = set(week_data.columns)
        if existing_columns & new_columns:
            raise ValueError("Duplicate columns detected after renaming.")

        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True)

    return transformed_df


def plot_district_graph(df, district_name, benchmark_brands, desired_diff, week_names, diff_week, download_pdf=False):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    
    plt.figure(figsize=(12, 8))
    district_df = df[df["Dist Name"] == district_name]
    price_diffs = []
    for brand in brands:
        brand_prices = []
        for week_name in week_names:
            column_name = f"{brand} ({week_name})"
            if column_name in district_df.columns:
                price = district_df[column_name].iloc[0]
                brand_prices.append(price)
            else:
                brand_prices.append(np.nan)
        valid_prices = [p for p in brand_prices if not np.isnan(p)]
        if len(valid_prices) > diff_week:
            price_diff = valid_prices[-1] - valid_prices[diff_week]
        else:
            price_diff = np.nan
        price_diffs.append(price_diff)
        plt.plot(week_names, brand_prices, marker='o', linestyle='-', label=f"{brand} ({price_diff:.0f})")
        for week, price in zip(week_names, brand_prices):
            if not np.isnan(price):
                plt.text(week, price, str(round(price)), fontsize=10)
    
    plt.grid(True, linestyle='--', alpha=0.7)
    plt.xlabel('Month/Week', weight='bold')
    plt.ylabel('Whole Sale Price (in Rs.)', weight='bold')
    region_name = district_df['REGION'].iloc[0]
    
    plt.text(0.5, 1.05, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)
    plt.title(f"{district_name} - Brands Price Trend", weight='bold', fontsize=18)
    
    plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=3, prop={'weight': 'bold'})
    plt.tight_layout()

    text_str = ''
    if benchmark_brands:
        brand_texts = []
        max_left_length = 0
        for benchmark_brand in benchmark_brands:
            jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
            benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
            actual_diff = np.nan
            if jklc_prices and benchmark_prices:
                for i in range(len(jklc_prices) - 1, -1, -1):
                    if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                        actual_diff = jklc_prices[i] - benchmark_prices[i]
                        break
            desired_diff_str = f" ({desired_diff[benchmark_brand]:.0f} Rs.)" if benchmark_brand in desired_diff and desired_diff[benchmark_brand] is not None else ""
            brand_text = [f"Benchmark Brand: {benchmark_brand}{desired_diff_str}", f"Actual Diff: {actual_diff:+.2f} Rs."]
            brand_texts.append(brand_text)
            max_left_length = max(max_left_length, len(brand_text[0]))
        num_brands = len(brand_texts)
        if num_brands == 1:
            text_str = "\n".join(brand_texts[0])
        elif num_brands > 1:
            half_num_brands = num_brands // 2
            left_side = brand_texts[:half_num_brands]
            right_side = brand_texts[half_num_brands:]
            lines = []
            for i in range(2):
                left_text = left_side[0][i] if i < len(left_side[0]) else ""
                right_text = right_side[0][i] if i < len(right_side[0]) else ""
                lines.append(f"{left_text.ljust(max_left_length)} \u2502 {right_text.rjust(max_left_length)}")
            text_str = "\n".join(lines)
    plt.text(0.5, -0.3, text_str, weight='bold', ha='center', va='center', transform=plt.gca().transAxes, bbox=dict(facecolor='white', edgecolor='black', boxstyle='round,pad=0.5'))
    plt.subplots_adjust(bottom=0.3)
    
    return plt.gcf()

def process_file():
    st.session_state.file_processed = True

def main():
    st.title("District Price Trend Analysis")

    col1, col2 = st.columns(2)

    with col1:
        uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

    if uploaded_file and not st.session_state.file_processed:
        try:
            file_content = uploaded_file.read()
            wb = openpyxl.load_workbook(BytesIO(file_content))
            ws = wb.active
            hidden_cols = [idx for idx, col in enumerate(ws.column_dimensions, 1) if ws.column_dimensions[col].hidden]

            st.session_state.df = pd.read_excel(BytesIO(file_content), skiprows=2)
            st.session_state.df.drop(st.session_state.df.columns[hidden_cols], axis=1, inplace=True)

            brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            brand_columns = [col for col in st.session_state.df.columns if any(brand in col for brand in brands)]

            num_weeks = len(brand_columns) // len(brands)
            
            with col2:
                st.subheader("Enter Week Names")
                for i in range(num_weeks):
                    st.session_state.week_names_input.append(st.text_input(f'Week {i+1}', key=f'week_{i}'))
                if st.button('Confirm Week Names', key='confirm_weeks'):
                    process_file()

        except Exception as e:
            st.error(f"Error processing file: {e}")

    if st.session_state.file_processed:
        st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.session_state.diff_week = st.slider("Select Week for Difference Calculation", 
                                                   min_value=0, 
                                                   max_value=len(st.session_state.week_names_input) - 1, 
                                                   value=st.session_state.diff_week, 
                                                   key="diff_week_slider")
        
        with col2:
            zone_names = st.session_state.df["Zone"].unique().tolist()
            selected_zone = st.selectbox("Select Zone", zone_names, key="zone_select")
        
        filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
        
        with col3:
            region_names = filtered_df["REGION"].unique().tolist()
            selected_region = st.selectbox("Select Region", region_names, key="region_select")
        
        filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
        
        district_names = filtered_df["Dist Name"].unique().tolist()
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
        
        brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
        benchmark_brands = [brand for brand in brands if brand != 'JKLC']
        
        for district in selected_districts:
            st.subheader(f"Settings for {district}")
            col1, col2 = st.columns(2)
            
            with col1:
                st.session_state.district_benchmark_brands[district] = st.multiselect(
                    f"Select Benchmark Brands for {district}", 
                    benchmark_brands, 
                    key=f"benchmark_select_{district}"
                )
            
            with col2:
                for benchmark_brand in st.session_state.district_benchmark_brands[district]:
                    st.session_state.desired_diff_input[f"{district}_{benchmark_brand}"] = st.number_input(
                        f"Desired Difference for {benchmark_brand} in {district}", 
                        min_value=-100.00, 
                        step=0.1, 
                        format="%.2f", 
                        key=f"{district}_{benchmark_brand}"
                    )
        
        download_pdf = st.checkbox("Download Plots as PDF")
        
        if st.button('Generate Plots', key='generate_plots'):
            pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf") if download_pdf else None
            
            for district in selected_districts:
                fig = plot_district_graph(
                    filtered_df, 
                    district, 
                    st.session_state.district_benchmark_brands[district],
                    {brand: st.session_state.desired_diff_input[f"{district}_{brand}"] for brand in st.session_state.district_benchmark_brands[district]},
                    st.session_state.week_names_input,
                    st.session_state.diff_week,
                    download_pdf
                )
                
                st.pyplot(fig)
                if download_pdf:
                    pdf.savefig(fig)
                plt.close(fig)
                
                buf = BytesIO()
                fig.savefig(buf, format='png', bbox_inches='tight')
                buf.seek(0)
                b64_data = base64.b64encode(buf.getvalue()).decode()
                st.markdown(f'<a download="district_plot_{district}.png" href="data:image/png;base64,{b64_data}">Download Plot for {district} as PNG</a>', unsafe_allow_html=True)
            
            if download_pdf:
                pdf.close()
                with open("district_plots.pdf", "rb") as f:
                    pdf_data = f.read()
                b64_pdf = base64.b64encode(pdf_data).decode()
                st.markdown(f'<a download="district_plots.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>', unsafe_allow_html=True)

if __name__ == "__main__":
    main()
