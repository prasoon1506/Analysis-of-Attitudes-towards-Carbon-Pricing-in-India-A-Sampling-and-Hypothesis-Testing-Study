import streamlit as st
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import matplotlib.backends.backend_pdf

st.set_page_config(page_title="WSP Analysis", layout="wide")

# Custom CSS for the entire app
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #f5f7fa 0%, #c3cfe2 100%);
    }
    .main .block-container {
        padding: 2rem;
        background: rgba(255, 255, 255, 0.9);
        border-radius: 15px;
        box-shadow: 0 8px 32px 0 rgba(31, 38, 135, 0.37);
    }
    h1 {
        color: #2c3e50;
        text-align: center;
        padding: 1.5rem;
        background: rgba(255, 255, 255, 0.95);
        border-radius: 10px;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stSelectbox, .stMultiSelect {
        background: white;
        border-radius: 8px;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .stButton > button {
        width: 100%;
        border-radius: 8px;
        background-color: #3498db;
        color: white;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton > button:hover {
        background-color: #2980b9;
        transform: translateY(-2px);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .stSlider > div > div > div {
        background-color: #3498db;
    }
    .stCheckbox > label {
        color: #2c3e50;
        font-weight: 500;
    }
    .stSubheader {
        color: #34495e;
        background: rgba(255, 255, 255, 0.9);
        padding: 0.8rem;
        border-radius: 8px;
        margin-top: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .uploadedFile {
        background-color: #e8f0fe;
        border-radius: 8px;
        padding: 1rem;
        margin-bottom: 1.5rem;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .dataframe {
        font-size: 0.8em;
    }
    .dataframe thead tr th {
        background-color: #3498db;
        color: white;
    }
    .dataframe tbody tr:nth-child(even) {
        background-color: #f2f2f2;
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

def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    
    # Region name replacements
    region_replacements = {
        '12_Madhya Pradesh(west)': 'Madhya Pradesh(West)',
        '20_Rajasthan': 'Rajasthan', '50_Rajasthan III': 'Rajasthan', '80_Rajasthan II': 'Rajasthan',
        '33_Chhattisgarh(2)': 'Chhattisgarh', '38_Chhattisgarh(3)': 'Chhattisgarh', '39_Chhattisgarh(1)': 'Chhattisgarh',
        '07_Haryana 1': 'Haryana', '07_Haryana 2': 'Haryana',
        '06_Gujarat 1': 'Gujarat', '66_Gujarat 2': 'Gujarat', '67_Gujarat 3': 'Gujarat', '68_Gujarat 4': 'Gujarat', '69_Gujarat 5': 'Gujarat',
        '13_Maharashtra': 'Maharashtra(West)',
        '24_Uttar Pradesh': 'Uttar Pradesh(West)',
        '35_Uttarakhand': 'Uttarakhand',
        '83_UP East Varanasi Region': 'Varanasi',
        '83_UP East Lucknow Region': 'Lucknow',
        '30_Delhi': 'Delhi',
        '19_Punjab': 'Punjab',
        '09_Jammu&Kashmir': 'Jammu&Kashmir',
        '08_Himachal Pradesh': 'Himachal Pradesh',
        '82_Maharashtra(East)': 'Maharashtra(East)',
        '81_Madhya Pradesh': 'Madhya Pradesh(East)',
        '34_Jharkhand': 'Jharkhand',
        '18_ODISHA': 'Odisha',
        '04_Bihar': 'Bihar',
        '27_Chandigarh': 'Chandigarh',
        '82_Maharashtra (East)': 'Maharashtra(East)',
        '25_West Bengal': 'West Bengal'
    }
    
    transformed_df['REGION'] = transformed_df['REGION'].replace(region_replacements)
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Delhi', 'Haryana', 'Punjab'], 'North-I')
    transformed_df['REGION'] = transformed_df['REGION'].replace(['Uttar Pradesh(West)','Uttarakhand'], 'North-II')
    
    zone_replacements = {
        'EZ_East Zone': 'East Zone',
        'CZ_Central Zone': 'Central Zone',
        'NZ_North Zone': 'North Zone',
        'UPEZ_UP East Zone': 'UP East Zone',
        'upWZ_up West Zone': 'UP West Zone',
        'WZ_West Zone': 'West Zone'
    }
    transformed_df['Zone'] = transformed_df['Zone'].replace(zone_replacements)
    
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    
    for i in range(num_weeks):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        
        # Use a unique suffix for each merge operation
        suffix = f'_{i}'
        transformed_df = pd.merge(transformed_df, week_data, left_index=True, right_index=True, suffixes=('', suffix))
    
    # Remove any columns with suffixes (duplicates)
    transformed_df = transformed_df.loc[:, ~transformed_df.columns.str.contains('_\d+$')]
    
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands_dict, desired_diff_dict, week_names, diff_week, download_pdf=False):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    if download_pdf:
        pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf")
    
    for i, district_name in enumerate(district_names):
        fig,ax=plt.subplots(figsize=(10, 8))
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
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        plt.ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        if i == 0:
            plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        else:
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()

        text_str = ''
        if district_name in benchmark_brands_dict:
            brand_texts = []
            max_left_length = 0
            for benchmark_brand in benchmark_brands_dict[district_name]:
                jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
                benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
                actual_diff = np.nan
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices) - 1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break
                desired_diff_str = f" ({desired_diff_dict[district_name][benchmark_brand]:.0f} Rs.)" if district_name in desired_diff_dict and benchmark_brand in desired_diff_dict[district_name] else ""
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
        plt.subplots_adjust(bottom=0.25)
        if download_pdf:
            pdf.savefig(fig, bbox_inches='tight')
        st.pyplot(fig)
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        b64_data = base64.b64encode(buf.getvalue()).decode()
        st.markdown(f'<a download="district_plot_{district_name}.png" href="data:image/png;base64,{b64_data}">Download Plot as PNG</a>', unsafe_allow_html=True)
        plt.close()
    
    if download_pdf:
        pdf.close()
        with open("district_plots.pdf", "rb") as f:
            pdf_data = f.read()
        b64_pdf = base64.b64encode(pdf_data).decode()
        st.markdown(f'<a download="{region_name}.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>', unsafe_allow_html=True)

def process_file():
    st.session_state.file_processed = True

def update_week_name(index):
    def callback():
        if index < len(st.session_state.week_names_input):
            st.session_state.week_names_input[index] = st.session_state[f'week_{index}']
        else:
            st.warning(f"Attempted to update week {index + 1}, but only {len(st.session_state.week_names_input)} weeks are available.")
        if all(st.session_state.week_names_input):
            process_file()
    return callback

def home():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: #3366cc;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)
    st.markdown('<div class="title"><span>Welcome to the WSP Analysis Dashboard</span></div>', unsafe_allow_html=True)

    
    st.markdown("""
    This app helps you analyze Whole Sale Price (WSP) data for various brands across different regions and districts.

    ## How to use this app:

    1. **Navigate to the WSP Analysis Dashboard tab** using the dropdown menu at the top of the sidebar.

    2. **Upload your Excel file** containing the WSP data.

    3. **Enter the week names** for each column in your data.

    4. **Select your analysis settings**:
        - Choose the zone and region you want to analyze
        - Select one or more districts
        - Set the week for difference calculation
        - Choose whether to download plots as PDF

    5. **Set benchmark brands and desired differences**:
        - You can set the same benchmarks for all districts or customize for each
        - For each benchmark brand, set the desired price difference

    6. **Generate plots** by clicking the 'Generate Plots' button

    7. **View and download** the generated plots

    Remember, you can always return to this page for a refresher on how to use the app.

    Happy analyzing!
    """)

def wsp_analysis_dashboard():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: #3366cc;
        text-align: center;
        padding: 20px;
        border-radius: 10px;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        margin-bottom: 30px;
        font-family: 'Arial', sans-serif;
    }
    .title span {
        background: linear-gradient(45deg, #3366cc, #6699ff);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    </style>
    """, unsafe_allow_html=True)

    # Display the stylized title
    st.markdown('<div class="title"><span>WSP Analysis Dashboard</span></div>', unsafe_allow_html=True)


    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        st.markdown(f'<div class="uploadedFile">File uploaded: {uploaded_file.name}</div>', unsafe_allow_html=True)

    if uploaded_file and not st.session_state.file_processed:
        try:
            file_content = uploaded_file.read()
            wb = openpyxl.load_workbook(BytesIO(file_content))
            ws = wb.active
            
            hidden_cols = [idx for idx, col in enumerate(ws.column_dimensions, 1) if ws.column_dimensions[col].hidden]
            
            st.session_state.df = pd.read_excel(BytesIO(file_content), skiprows=2)
            
            if st.session_state.df.empty:
                st.error("The uploaded file resulted in an empty dataframe. Please check the file content.")
            else:
                st.session_state.df.drop(st.session_state.df.columns[hidden_cols], axis=1, inplace=True)

                brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                brand_columns = [col for col in st.session_state.df.columns if any(brand in col for brand in brands)]

                num_weeks = len(brand_columns) // len(brands)
                
                if num_weeks > 0:
                    st.markdown("### Enter Week Names")
                    num_columns = max(1, num_weeks)
                    week_cols = st.columns(num_columns)
                    
                    if 'week_names_input' not in st.session_state or len(st.session_state.week_names_input) != num_weeks:
                        st.session_state.week_names_input = [''] * num_weeks
                    
                    for i in range(num_weeks):
                        with week_cols[i % num_columns]:
                            st.text_input(
                                f'Week {i+1}', 
                                value=st.session_state.week_names_input[i] if i < len(st.session_state.week_names_input) else '',
                                key=f'week_{i}',
                                on_change=update_week_name(i)
                            )
                else:
                    st.warning("No weeks detected in the uploaded file. Please check the file content.")
                    st.session_state.week_names_input = []

        except Exception as e:
            st.error(f"Error processing file: {e}")
            st.exception(e)

    if st.session_state.file_processed:
        st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)
        
        st.markdown("### Analysis Settings")
        
        st.session_state.diff_week = st.slider("Select Week for Difference Calculation", 
                                               min_value=0, 
                                               max_value=len(st.session_state.week_names_input) - 1, 
                                               value=st.session_state.diff_week, 
                                               key="diff_week_slider")
        
        download_pdf = st.checkbox("Download Plots as PDF")
        
        col1, col2 = st.columns(2)
        with col1:
            zone_names = st.session_state.df["Zone"].unique().tolist()
            selected_zone = st.selectbox("Select Zone", zone_names, key="zone_select")
        with col2:
            filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
            region_names = filtered_df["REGION"].unique().tolist()
            selected_region = st.selectbox("Select Region", region_names, key="region_select")
        
        filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
        district_names = filtered_df["Dist Name"].unique().tolist()
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")

        brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
        benchmark_brands = [brand for brand in brands if brand != 'JKLC']
        
        benchmark_brands_dict = {}
        desired_diff_dict = {}
        
        if selected_districts:
            st.markdown("### Benchmark Settings")
            use_same_benchmarks = st.checkbox("Use same benchmarks for all districts", value=True)
            
            if use_same_benchmarks:
                selected_benchmarks = st.multiselect("Select Benchmark Brands for all districts", benchmark_brands, key="unified_benchmark_select")
                for district in selected_districts:
                    benchmark_brands_dict[district] = selected_benchmarks
                    desired_diff_dict[district] = {}
    
                if selected_benchmarks:
                    st.markdown("#### Desired Differences")
                    num_cols = min(len(selected_benchmarks), 3)
                    diff_cols = st.columns(num_cols)
                    for i, brand in enumerate(selected_benchmarks):
                        with diff_cols[i % num_cols]:
                            value = st.number_input(
                                f"{brand}",
                                min_value=-100.00,
                                step=0.1,
                                format="%.2f",
                                key=f"unified_{brand}"
                            )
                            for district in selected_districts:
                                desired_diff_dict[district][brand] = value
                else:
                    st.warning("Please select at least one benchmark brand.")
            else:
                for district in selected_districts:
                    st.subheader(f"Settings for {district}")
                    benchmark_brands_dict[district] = st.multiselect(
                        f"Select Benchmark Brands for {district}",
                        benchmark_brands,
                        key=f"benchmark_select_{district}"
                    )
                    desired_diff_dict[district] = {}
                    
                    if benchmark_brands_dict[district]:
                        num_cols = min(len(benchmark_brands_dict[district]), 3)
                        diff_cols = st.columns(num_cols)
                        for i, brand in enumerate(benchmark_brands_dict[district]):
                            with diff_cols[i % num_cols]:
                                desired_diff_dict[district][brand] = st.number_input(
                                    f"{brand}",
                                    min_value=-100.00,
                                    step=0.1,
                                    format="%.2f",
                                    key=f"{district}_{brand}"
                                )
                    else:
                        st.warning(f"No benchmark brands selected for {district}.")
        
        st.markdown("### Generate Analysis")
        
        if st.button('Generate Plots', key='generate_plots', use_container_width=True):
            with st.spinner('Generating plots...'):
                plot_district_graph(filtered_df, selected_districts, benchmark_brands_dict, 
                                    desired_diff_dict, 
                                    st.session_state.week_names_input, 
                                    st.session_state.diff_week, 
                                    download_pdf)
                st.success('Plots generated successfully!')

def main():
    st.sidebar.title("Navigation")
    app_mode = st.sidebar.selectbox("Choose the app mode",
        ["Home", "WSP Analysis Dashboard"])
    
    if app_mode == "Home":
        home()
    elif app_mode == "WSP Analysis Dashboard":
        wsp_analysis_dashboard()

if __name__ == "__main__":
    main()
