import openpyxl
from IPython import get_ipython
from IPython.display import display
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from ipywidgets import Dropdown, interact, FileUpload, Button, HBox, VBox, Output, fixed, SelectMultiple, interactive, IntText, Text
from io import BytesIO
from IPython.display import display, FileLink, HTML, clear_output
import base64
from tqdm import tqdm
import xgboost as xgb
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
from scipy import stats
import matplotlib.backends.backend_pdf

df = None
desired_diff_input = {}
week_name_widgets = []  # Initialize week_name_widgets as an empty list
confirm_button = None   # Initialize desired_diff_input as a dictionary

def transform_data(df, week_names_input):
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    transformed_df = df[['Zone', 'REGION', 'Dist Code', 'Dist Name']].copy()
    brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]
    num_weeks = len(brand_columns) // len(brands)
    for i in tqdm(range(num_weeks), desc='Transforming data'):
        start_idx = i * len(brands)
        end_idx = (i + 1) * len(brands)
        week_data = df[brand_columns[start_idx:end_idx]]
        week_name = week_names_input[i]  # Use week name from user input
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


def plot_district_graph(df, district_names, benchmark_brands, desired_diff, week_names, download_pdf=False, diff_week=1):
    clear_output(wait=True)
    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    num_weeks = len(df.columns[4:]) // len(brands)
    if download_pdf:
        pdf = matplotlib.backends.backend_pdf.PdfPages("district_plots.pdf")
    for i, district_name in enumerate(district_names):
        plt.figure(figsize=(10, 8))
        district_df = df[df["Dist Name"] == district_name]
        price_diffs = []
        for brand in brands:
            brand_prices = []
            for week_name in week_names:  # Use week_names directly
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
            line, = plt.plot(week_names,  # Use week_names directly
                             brand_prices,
                             marker='o',
                             linestyle='-',
                             label=f"{brand} ({price_diff:.0f})")
            for week, price in zip(week_names, brand_prices):  # Use week_names directly
                if not np.isnan(price):
                    plt.text(week, price, str(round(price)), fontsize=10)
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        plt.ylabel('Whole Sale Price(in Rs.)', weight='bold')
        region_name = district_df['REGION'].iloc[0]
        
        # Add region name above the title only for the first district
        if i == 0:
            #plt.title(f"{region_name}\n{district_name} - Brands Price Trend", weight='bold',fontsize=16)
            plt.text(0.5, 1.1, region_name, ha='center', va='center', transform=plt.gca().transAxes, weight='bold', fontsize=16)  # Added region name using plt.text
            plt.title(f"{district_name} - Brands Price Trend", weight='bold') # Keep the original title without region name
        else:
            plt.title(f"{district_name} - Brands Price Trend", weight='bold')
        
        plt.legend(loc='upper center', bbox_to_anchor=(0.5, -0.15), ncol=6, prop={'weight': 'bold'})
        plt.tight_layout()

        text_str = ''
        if benchmark_brands:
            brand_texts = []
            max_left_length = 0  # Store text for each brand separately
            for benchmark_brand in benchmark_brands:
                jklc_prices = [district_df[f"JKLC ({week})"].iloc[0] for week in week_names if f"JKLC ({week})" in district_df.columns]
                benchmark_prices = [district_df[f"{benchmark_brand} ({week})"].iloc[0] for week in week_names if f"{benchmark_brand} ({week})" in district_df.columns]
                actual_diff = np.nan  # Initialize actual_diff with NaN
                if jklc_prices and benchmark_prices:
                    for i in range(len(jklc_prices) - 1, -1, -1):
                        if not np.isnan(jklc_prices[i]) and not np.isnan(benchmark_prices[i]):
                            actual_diff = jklc_prices[i] - benchmark_prices[i]
                            break
                desired_diff_str = f" ({desired_diff[benchmark_brand].value:.0f} Rs.)" if benchmark_brand in desired_diff and desired_diff[benchmark_brand].value is not None else ""
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
            pdf.savefig()
        buf = BytesIO()
        plt.savefig(buf, format='png', bbox_inches='tight')
        buf.seek(0)
        b64_data = base64.b64encode(buf.getvalue()).decode()
        display(
            HTML(
                f'<a download="district_plot_{district_name}.png" href="data:image/png;base64,{b64_data}">Download Plot as PNG</a>'
            ))
        plt.show()    
    if download_pdf:
       pdf.close()
       with open("district_plots.pdf", "rb") as f:
           pdf_data = f.read()
       b64_pdf = base64.b64encode(pdf_data).decode()
       display(HTML(f'<a download="{region_name}.pdf" href="data:application/pdf;base64,{b64_pdf}">Download All Plots as PDF</a>'))   

def on_button_click(change):
    global week_name_widgets, confirm_button # Access global variables
    for widget in week_name_widgets:
        widget.close() # Close existing week name widgets
    if confirm_button:
        confirm_button.close() # Close existing confirm button
    uploaded_file = upload_button.value
    if uploaded_file:
        try:
            file_name = list(uploaded_file.keys())[0]
            file_content = uploaded_file[file_name]['content']
            wb = openpyxl.load_workbook(BytesIO(file_content))
            ws = wb.active
            hidden_cols = [idx for idx, col in enumerate(ws.column_dimensions, 1) if ws.column_dimensions[col].hidden]

            global df
            df = pd.read_excel(BytesIO(file_content), skiprows=2)
            df.drop(df.columns[hidden_cols], axis=1, inplace=True)

            
            brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            brand_columns = [col for col in df.columns if any(brand in col for brand in brands)]

            # Get week names using Text widgets
            num_weeks = len(brand_columns) // 6  # Calculate num_weeks using brand_columns
            week_name_widgets = [Text(description=f'Week {i+1}:') for i in range(num_weeks)]
            display(*week_name_widgets)

            def get_week_names(button):
                week_names_input = [widget.value for widget in week_name_widgets]
                global df
                df = transform_data(df, week_names_input)
                zone_names = df["Zone"].unique().tolist()
                zone_dropdown.options = zone_names
                with output:
                    output.clear_output(wait=True)
                    print(f"Uploaded file: {file_name}")
                    create_interactive_plot(df, week_names_input)

            confirm_button = Button(description="Confirm Week Names")
            confirm_button.on_click(get_week_names)
            display(confirm_button)

        except Exception as e:
            with output:
                output.clear_output()
                print(f"Error reading file: {e}. Please ensure it is a valid Excel file.")

def on_zone_change(change):
    if change['type'] == 'change' and change['name'] == 'value':
        selected_zone = change['new']
        global df
        filtered_df = df[df["Zone"] == selected_zone]
        region_names = filtered_df["REGION"].unique().tolist()
        region_dropdown.options = region_names
        district_dropdown.options = []

def on_region_change(change):
    if change['type'] == 'change' and change['name'] == 'value':
        selected_region = change['new']
        global df
        filtered_df = df[df["REGION"] == selected_region]
        district_names = filtered_df["Dist Name"].unique().tolist()
        district_dropdown.options = district_names

def on_district_change(change):
    if change['type'] == 'change' and change['name'] == 'value':
        selected_districts = change['new']
        global df, desired_diff_input
        if selected_districts:
            all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
            benchmark_dropdown.options = benchmark_brands
from ipywidgets import IntSlider
def create_interactive_plot(df, week_names_input):
    global desired_diff_input
    all_brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
    benchmark_brands = [brand for brand in all_brands if brand != 'JKLC']
    desired_diff_input = {brand: IntText(value=None, description=f'Desired Diff for {brand}:') for brand in benchmark_brands}
    desired_diff_copy = {brand: IntText(value=diff_input.value, description=diff_input.description) for brand, diff_input in desired_diff_input.items()}
    num_weeks = len(week_names_input)
    diff_week_slider = IntSlider(value=0, min=0, max=num_weeks-1, step=1, description='Diff Week:')
    w = interactive(plot_district_graph,
                    {'manual': True},
                    df=fixed(df),
                    district_names=district_dropdown,
                    benchmark_brands=benchmark_dropdown,
                    desired_diff=fixed(desired_diff_copy),
                    week_names=fixed(week_names_input),
                    download_pdf=False,diff_week=diff_week_slider)
    for brand, diff_input in desired_diff_copy.items():
        w.children += (diff_input,)

    download_pdf_button = Button(description='Download PDF')

    def on_download_pdf_button_clicked(b):
        w.kwargs['download_pdf'] = True
        w.update()

    download_pdf_button.on_click(on_download_pdf_button_clicked)
    w.children[-7].description = "Run Interact"
    display(w)

upload_button = FileUpload(accept='.xlsx', description="Upload Excel")
select_button = Button(description="Select File")
output = Output()

select_button.on_click(on_button_click)

zone_dropdown = Dropdown(options=[], description="Select Zone")
region_dropdown = Dropdown(options=[], description="Select Region")
district_dropdown = SelectMultiple(options=[], description="Select District")
benchmark_dropdown = SelectMultiple(options=[], description="Select Benchmark Brands")

zone_dropdown.observe(on_zone_change, names='value')
region_dropdown.observe(on_region_change, names='value')
district_dropdown.observe(on_district_change, names='value')

file_upload_area = VBox([upload_button, select_button, output])
selection_area = HBox([zone_dropdown, region_dropdown])
main_area = VBox([file_upload_area, selection_area])
display(main_area)
