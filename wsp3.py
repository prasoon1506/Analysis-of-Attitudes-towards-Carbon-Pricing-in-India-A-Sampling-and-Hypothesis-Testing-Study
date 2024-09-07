import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import base64
from io import BytesIO
from tqdm import tqdm
import matplotlib.backends.backend_pdf
import openpyxl

# Initialize session state variables
if 'df' not in st.session_state:
    st.session_state.df = None
if 'week_names' not in st.session_state:
    st.session_state.week_names = []

def read_excel_excluding_hidden_columns(file):
    # Load the workbook and select the active worksheet
    wb = openpyxl.load_workbook(file, read_only=True)
    ws = wb.active

    # Get the indices of hidden columns
    hidden_cols = [i for i, col in enumerate(ws.column_dimensions.values(), 1) if col.hidden]

    # Read the Excel file with pandas, skipping hidden columns
    df = pd.read_excel(file, skiprows=2, usecols=lambda x: x+1 not in hidden_cols)

    return df

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
        week_data = week_data.rename(columns={
            col: f"{brand} ({week_name})"
            for brand, col in zip(brands, week_data.columns)
        })
        week_data.replace(0, np.nan, inplace=True)
        transformed_df = pd.merge(transformed_df,
                                  week_data,
                                  left_index=True,
                                  right_index=True)
    return transformed_df

def plot_district_graph(df, district_names, benchmark_brands, desired_diff, week_names, diff_week=1):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    pdf_buffer = BytesIO()
    pdf = matplotlib.backends.backend_pdf.PdfPages(pdf_buffer)
    
    for district_name in district_names:
        fig, ax = plt.subplots(figsize=(10, 8))
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
            line, = ax.plot(week_names,
                            brand_prices,
                            marker='o',
                            linestyle='-',
                            label=f"{brand} ({price_diff:.0f})")
            for week, price in zip(week_names, brand_prices):
                if not np.isnan(price):
                    ax.text(week, price, str(round(price)), fontsize=10)
        
        ax.grid(False)
        ax.set_xlabel('Month/Week', weight='bold')
        ax.set_ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        ax.text(0.5, 1.1, region_name, ha='center', va='center', transform=ax.transAxes, weight='bold', fontsize=16)
        ax.set_title(f"{district_name} - Brands Price Trend", weight='bold')
        
        ax.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
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
                desired_diff_str = f" ({desired_diff[benchmark_brand]:.0f} Rs.)" if benchmark_brand in desired_diff else ""
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
        
        pdf.savefig(fig)
        st.pyplot(fig)
        plt.close(fig)
    
    pdf.close()
    pdf_buffer.seek(0)
    return pdf_buffer

def main():
    st.title("WSP Analysis App")

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            df = read_excel_excluding_hidden_columns(uploaded_file)
            st.session_state.df = df

            brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
            num_weeks = len(brand_columns) // 6

            week_names = []
            for i in range(num_weeks):
                week_name = st.text_input(f"Week {i+1} name:", key=f"week_{i}")
                week_names.append(week_name)

            if st.button("Confirm Week Names"):
                st.session_state.week_names = week_names
                st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names)
                st.success("Data transformed successfully!")

            if st.session_state.df is not None:
                zone_names = st.session_state.df["Zone"].unique().tolist()
                selected_zone = st.selectbox("Select Zone", zone_names)

                filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
                region_names = filtered_df["REGION"].unique().tolist()
                selected_region = st.selectbox("Select Region", region_names)

                filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
                district_names = filtered_df["Dist Name"].unique().tolist()
                selected_districts = st.multiselect("Select Districts", district_names)

                all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
                selected_benchmark_brands = st.multiselect("Select Benchmark Brands", benchmark_brands)

                desired_diff = {}
                for brand in selected_benchmark_brands:
                    desired_diff[brand] = st.number_input(f"Desired Diff for {brand}", value=0)

                diff_week = st.slider("Diff Week", min_value=0, max_value=len(st.session_state.week_names)-1, value=1)

                if st.button("Generate Plots"):
                    pdf_buffer = plot_district_graph(filtered_df, selected_districts, selected_benchmark_brands, desired_diff, st.session_state.week_names, diff_week)
                    
                    st.download_button(
                        label="Download PDF",
                        data=pdf_buffer,
                        file_name="district_plots.pdf",
                        mime="application/pdf"
                    )

        except Exception as e:
            st.error(f"Error reading file: {e}. Please ensure it is a valid Excel file.")

if __name__ == "__main__":
    main()
