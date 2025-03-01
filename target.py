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
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error
import xgboost as xgb
import lightgbm as lgb
from sklearn.ensemble import VotingRegressor
from sklearn.model_selection import GridSearchCV
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    regions = df['Zone'].unique().tolist()
    brands = df['Brand'].unique().tolist()
    return df, regions, brands
def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
from sklearn.ensemble import VotingRegressor
@st.cache_resource
def train_advanced_model(X_train, y_train):
    # Feature scaling
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    models = {'XGBoost': xgb.XGBRegressor(random_state=42),'LightGBM': lgb.LGBMRegressor(random_state=42),'RandomForest': RandomForestRegressor(random_state=42)}
    best_model = None
    best_score = float('-inf')
    for name, model in models.items():
        if X_train.shape[0] < 5:
            model.fit(X_train_scaled, y_train)
            score = model.score(X_train_scaled, y_train)
        else:
            scores = cross_val_score(model, X_train_scaled, y_train, cv=min(5, X_train.shape[0]))
            score = np.mean(scores)
        if score > best_score:
            best_score = score
            best_model = model
    best_model.fit(X_train_scaled, y_train)
    return best_model, scaler
def predict_and_visualize(df, region, brand):
    try:
        region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)].copy()
        if len(region_data) > 0:
            months = ['Apr', 'May', 'June', 'July', 'Aug']
            region_data['YoY_Growth'] = (region_data['Monthly Achievement(Aug)'] - region_data['Total Aug 2023']) / region_data['Total Aug 2023']
            X = pd.DataFrame({'Month_Tgt': [region_data[f'Month Tgt ({month})'].iloc[-1] for month in months],'Achievement': [region_data[f'Monthly Achievement({month})'].iloc[-1] for month in months],'YoY_Growth': [region_data['YoY_Growth'].iloc[-1]] * len(months),'Last_Year_Sep_Sales': [region_data['Total Sep 2023'].iloc[-1]] * len(months),'Month_Number': range(4, 9)})
            y = region_data[[f'Monthly Achievement({month})' for month in months]].values.ravel()
            if X.shape[0] < 5:
                X_train, X_val, y_train, y_val = X, X, y, y
            else:
                X_train, X_val, y_train, y_val = train_test_split(X, y, test_size=0.2, random_state=42)
            model, scaler = train_advanced_model(X_train, y_train)
            sept_target = region_data['Month Tgt (Sep)'].iloc[-1]
            sept_data = pd.DataFrame({'Month_Tgt': [sept_target],'Achievement': [region_data['Monthly Achievement(Aug)'].iloc[-1]],'YoY_Growth': [region_data['YoY_Growth'].iloc[-1]],'Last_Year_Sep_Sales': [region_data['Total Sep 2023'].iloc[-1]],'Month_Number': [9]  # September is 9})
            sept_data_scaled = scaler.transform(sept_data)
            sept_prediction = model.predict(sept_data_scaled)[0]
            y_val_pred = model.predict(scaler.transform(X_val))
            rmse = np.sqrt(mean_squared_error(y_val, y_val_pred))
            confidence = 0.95
            degrees_of_freedom = len(y_train) - X_train.shape[1] - 1
            t_value = stats.t.ppf((1 + confidence) / 2, degrees_of_freedom)
            prediction_std = np.std(y_val - y_val_pred)
            margin_of_error = t_value * prediction_std / np.sqrt(len(y_val))
            lower_bound = max(0, sept_prediction - margin_of_error)
            upper_bound = sept_prediction + margin_of_error
            sept_achievement = sept_prediction
            lower_achievement = lower_bound
            upper_achievement = upper_bound
            fig = create_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse)
            return fig, sept_achievement, lower_achievement, upper_achievement, rmse
        else:
            return None, None, None, None, None
    except Exception as e:
        st.error(f"Error in predict_and_visualize: {str(e)}")
        raise
def create_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse):
    fig = plt.figure(figsize=(20, 28))
    gs = fig.add_gridspec(8, 2, height_ratios=[0.5, 0.5, 0.5, 3, 1, 2, 1,1])
    ax_region = fig.add_subplot(gs[0, :])
    ax_region.axis('off')
    ax_region.text(0.5, 0.5, f'{region}({brand})', fontsize=28, fontweight='bold', ha='center', va='center')
    ax_current = fig.add_subplot(gs[1, :])
    ax_current.axis('off')
    current_data = [['Total Sales\nTill Now', 'Commitment\nfor Today', 'Asking\nfor Today', 'Yesterday\nSales', 'Yesterday\nCommitment'],[f"{region_data['Till Yesterday Total Sales'].iloc[-1]:.0f}",f"{region_data['Commitment for Today'].iloc[-1]:.0f}",f"{region_data['Asking for Today'].iloc[-1]:.0f}",f"{region_data['Yesterday Sales'].iloc[-1]:.0f}",f"{region_data['Yesterday Commitment'].iloc[-1]:.0f}"]]
    current_table = ax_current.table(cellText=current_data[1:], colLabels=current_data[0], cellLoc='center', loc='center')
    current_table.auto_set_font_size(False)
    current_table.set_fontsize(10)
    current_table.scale(1, 1.7)
    for (row, col), cell in current_table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='black')
                    cell.set_facecolor('goldenrod')
                cell.set_edgecolor('brown')
    ax_table = fig.add_subplot(gs[2, :])
    ax_table.axis('off')
    table_data = [['Month Target\n(Sep)', 'Monthly Achievement\n(Aug)', 'Predicted Achievement\n(Sept)(using XGBoost Algorithm)', 'CI', 'RMSE'],[f"{sept_target:.2f}", f"{region_data['Monthly Achievement(Aug)'].iloc[-1]:.2f}", f"{sept_achievement:.2f}", f"({lower_achievement:.2f}, {upper_achievement:.2f})", f"{rmse:.4f}"]]
    table = ax_table.table(cellText=table_data[1:], colLabels=table_data[0], cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.7)
    for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='black')
                    cell.set_facecolor('goldenrod')
                cell.set_edgecolor('brown')
    ax1 = fig.add_subplot(gs[3, :])
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
        ax1.text(i, (max(target, achievement)+min(target,achievement))/2, f'{percentage:.1f}%',ha='center', va='bottom', fontsize=10, color=color, fontweight='bold')
    ax1.errorbar(x[-1] + width/2, sept_achievement,yerr=[[sept_achievement - lower_achievement], [upper_achievement - sept_achievement]],fmt='o', color='darkred', capsize=5, capthick=2, elinewidth=2)
    ax2 = fig.add_subplot(gs[4, :])
    percent_achievements = [((ach / tgt) * 100) for ach, tgt in zip(all_achievements, all_targets)]
    ax2.plot(x, percent_achievements, marker='o', linestyle='-', color='purple')
    ax2.axhline(y=100, color='r', linestyle='--', alpha=0.7)
    ax2.set_xlabel('Month', fontsize=12, fontweight='bold')
    ax2.set_ylabel('% Achievement', fontsize=12, fontweight='bold')
    ax2.set_xticks(x)
    ax2.set_xticklabels(all_months)
    for i, pct in enumerate(percent_achievements):
        ax2.annotate(f'{pct:.1f}%', (i, pct), xytext=(0, 5), textcoords='offset points', ha='center', va='bottom', fontsize=8)
    ax3 = fig.add_subplot(gs[5, :])
    ax3.axis('off')
    current_year = 2024  # Assuming the current year is 2024
    last_year = 2023
    channel_data = [('Trade', region_data['Trade Aug'].iloc[-1], region_data['Trade Aug 2023'].iloc[-1]),('Premium', region_data['Premium Aug'].iloc[-1], region_data['Premium Aug 2023'].iloc[-1]),('Blended', region_data['Blended Aug'].iloc[-1], region_data['Blended Aug 2023'].iloc[-1])]
    monthly_achievement_aug = region_data['Monthly Achievement(Aug)'].iloc[-1]
    total_aug_current = region_data['Monthly Achievement(Aug)'].iloc[-1]
    total_aug_last = region_data['Total Aug 2023'].iloc[-1]
    ax3.text(0.2, 1, f'\nAugust {current_year} Sales Breakdown:-', fontsize=16, fontweight='bold', ha='center', va='center')
    def get_arrow(value):
        return '↑' if value > 0 else '↓' if value < 0 else '→'
    def get_color(value):
        return 'green' if value > 0 else 'red' if value < 0 else 'black'
    total_change = ((total_aug_current - total_aug_last) / total_aug_last) * 100
    arrow = get_arrow(total_change)
    color = get_color(total_change)
    ax3.text(0.21, 0.9, f"August 2024: {total_aug_current:.0f}", fontsize=14, fontweight='bold', ha='center')
    ax3.text(0.22, 0.85, f"vs August 2023: {total_aug_last:.0f} ({total_change:.1f}% {arrow})", fontsize=12, color=color, ha='center')
    for i, (channel, value_current, value_last) in enumerate(channel_data):
        percentage = (value_current / monthly_achievement_aug) * 100
        change = ((value_current - value_last) / value_last) * 100
        arrow = get_arrow(change)
        color = get_color(change)
        y_pos = 0.75 - i*0.25
        ax3.text(0.1, y_pos, f"{channel}:", fontsize=14, fontweight='bold')
        ax3.text(0.2, y_pos, f"{value_current:.0f} ({percentage:.1f}%)", fontsize=14)
        ax3.text(0.1, y_pos-0.05, f"vs Last Year: {value_last:.0f}", fontsize=12)
        ax3.text(0.2, y_pos-0.05, f"({change:.1f}% {arrow})", fontsize=12, color=color)
    ax4 = fig.add_subplot(gs[5, 1])
    region_type_data = [region_data['Green Aug'].iloc[-1],region_data['Yellow Aug'].iloc[-1],region_data['Red Aug'].iloc[-1],region_data['Unidentified Aug'].iloc[-1]]
    region_type_labels = ['Green', 'Yellow', 'Red', 'Unidentified']
    colors = ['green', 'yellow', 'red', 'gray']
    def make_autopct(values):
        def my_autopct(pct):
            total = sum(values)
            val = int(round(pct*total/100.0))
            return f'{pct:.1f}%\n({val:.0f})'
        return my_autopct
    ax4.pie(region_type_data, labels=region_type_labels, colors=colors,autopct=make_autopct(region_type_data), startangle=90)
    ax4.set_title('August 2024 Region Type Breakdown:-', fontsize=16, fontweight='bold')
    ax5 = fig.add_subplot(gs[6, :])
    ax5.axis('off')
    q3_table_data = [['Overall Requirement', 'Requirement in\nTrade Channel', 'Requirement in\nBlednded Product Category', 'Requirement for\nPremium Product'],[f"{region_data['Q3 2023'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Trade'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Blended'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Premium'].iloc[-1]:.0f}"]]
    q3_table = ax5.table(cellText=q3_table_data[1:], colLabels=q3_table_data[0], cellLoc='center', loc='center')
    q3_table.auto_set_font_size(False)
    q3_table.set_fontsize(10)
    q3_table.scale(1, 1.7)
    for (row, col), cell in q3_table.get_celld().items():
        if row == 0:
            cell.set_text_props(fontweight='bold', color='black')
            cell.set_facecolor('goldenrod')
        cell.set_edgecolor('brown')
    ax5.set_title('Quarterly Requirements for September 2024', fontsize=16, fontweight='bold')
    ax_insights = fig.add_subplot(gs[7, :])
    ax_insights.axis('off')
    yoy_growth = (region_data['Monthly Achievement(Aug)'].iloc[-1] - region_data['Total Aug 2023'].iloc[-1]) / region_data['Total Aug 2023'].iloc[-1] * 100
    last_year_sept = region_data['Total Sep 2023'].iloc[-1]
    predicted_growth = (sept_achievement - last_year_sept) / last_year_sept * 100
    ax_insights.text(0.1, 0.8, f"Year-over-Year Growth (Aug): {yoy_growth:.2f}%", fontsize=12, fontweight='bold')
    ax_insights.text(0.1, 0.6, f"Last Year September Sales: {last_year_sept:.0f}", fontsize=12, fontweight='bold')
    ax_insights.text(0.1, 0.4, f"Predicted Growth (Sep): {predicted_growth:.2f}%", fontsize=12, fontweight='bold')
    plt.tight_layout()
    return fig
def generate_combined_report(df, regions, brands):
    main_table_data = [['Region', 'Brand', 'Month Target\n(Sep)', 'Monthly Achievement\n(Aug)', 'Predicted\nAchievement(Sept)', 'CI', 'RMSE']]
    additional_table_data = [['Region', 'Brand', 'Till Yesterday\nTotal Sales', 'Commitment\nfor Today', 'Asking\nfor Today', 'Yesterday\nSales', 'Yesterday\nCommitment']]
    with ThreadPoolExecutor() as executor:
        futures = []
        for region in regions:
            for brand in brands:
                futures.append(executor.submit(predict_and_visualize, df, region, brand))
        valid_data = False
        for future, (region, brand) in zip(futures, [(r, b) for r in regions for b in brands]):
            try:
                _, sept_achievement, lower_achievement, upper_achievement, rmse = future.result()
                if sept_achievement is not None:
                    region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)]
                    if not region_data.empty:
                        sept_target = region_data['Month Tgt (Sep)'].iloc[-1]
                        aug_achievement = region_data['Monthly Achievement(Aug)'].iloc[-1]
                        main_table_data.append([region, brand, f"{sept_target:.0f}", f"{aug_achievement:.0f}",f"{sept_achievement:.0f}", f"({lower_achievement:.2f},\n{upper_achievement:.2f})", f"{rmse:.4f}"])
                        additional_table_data.append([region, brand, f"{region_data['Till Yesterday Total Sales'].iloc[-1]:.0f}",f"{region_data['Commitment for Today'].iloc[-1]:.0f}",f"{region_data['Asking for Today'].iloc[-1]:.0f}",f"{region_data['Yesterday Sales'].iloc[-1]:.0f}",f"{region_data['Yesterday Commitment'].iloc[-1]:.0f}"])
                        valid_data = True
                    else:
                        st.warning(f"No data available for {region} and {brand}")
            except Exception as e:
                st.warning(f"Error processing {region} and {brand}: {str(e)}")
    if valid_data:
        num_rows = len(main_table_data) + len(additional_table_data)
        fig_height = max(12, 2 + 0.5 * num_rows)
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, fig_height), gridspec_kw={'height_ratios': [1, 1.5]})
        fig.suptitle("", fontsize=16, fontweight='bold', y=0.98)
        def create_styled_table(ax, data, title):
            ax.axis('off')
            ax.set_title(title, fontsize=14, fontweight='bold', pad=20)
            table = ax.table(cellText=data[1:], colLabels=data[0], cellLoc='center', loc='center')
            table.auto_set_font_size(False)
            table.set_fontsize(8)
            table.scale(1, 1.5)
            for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='white')
                    cell.set_facecolor('#4CAF50')
                elif row % 2 == 0:
                    cell.set_facecolor('#f2f2f2')
                cell.set_edgecolor('white')
                cell.set_text_props(wrap=True)
            for i in range(len(data[0])):
                table.auto_set_column_width(i)
        create_styled_table(ax1, additional_table_data, "Current Month Sales Data")
        create_styled_table(ax2, main_table_data, "Sales Predictions")
        plt.tight_layout()
        pdf_buffer = BytesIO()
        fig.savefig(pdf_buffer, format='pdf', bbox_inches='tight')
        plt.close(fig)
        pdf_buffer.seek(0)
        return base64.b64encode(pdf_buffer.getvalue()).decode()
    else:
        st.warning("No valid data available for any region and brand combination.")
        return None
def main():
    st.set_page_config(page_title="Sales Prediction App", page_icon="📊", layout="wide")
    lottie_url = "https://assets5.lottiefiles.com/packages/lf20_V9t630.json"
    lottie_json = load_lottie_url(lottie_url)
    with st.sidebar:
        st_lottie(lottie_json, height=200)
        st.title("Navigation")
        page = st.radio("Go to", ["Home", "Predictions","XGBoost Explained", "About"])
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
                    buf = BytesIO()
                    fig.savefig(buf, format="pdf")
                    buf.seek(0)
                    b64 = base64.b64encode(buf.getvalue()).decode()
                    st.download_button(label="Download Individual PDF Report",data=buf,file_name=f"prediction_report_{region}_{brand}.pdf",mime="application/pdf")
                else:
                    st.error(f"No data available for {region} and {brand}")
            if st.button("Generate Combined Report"):
                     with st.spinner("Generating combined report..."):
                         combined_report_data = generate_combined_report(df, regions, brands)
                     if combined_report_data:
                          st.download_button(label="Download Combined PDF Report",data=base64.b64decode(combined_report_data),file_name="combined_prediction_report.pdf",mime="application/pdf")
                     else:
                        st.error("Unable to generate combined report. Please check the warnings above for more details.")

if __name__ == "__main__":
    main()
