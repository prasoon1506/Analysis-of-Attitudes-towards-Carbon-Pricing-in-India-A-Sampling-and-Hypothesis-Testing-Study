import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
import xgboost as xgb
from io import BytesIO
import base64
from concurrent.futures import ThreadPoolExecutor

# Improved model functions
def preprocess_data(df):
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']
    for month in months:
        df[f'Achievement({month})'] = df[f'Monthly Achievement({month})'] / df[f'Month Tgt ({month})']
        df[f'Target({month})'] = df[f'Month Tgt ({month})']

    df['PrevYearSep'] = df['Total Sep 2023']
    df['PrevYearOct'] = df['Total Oct 2023']
    df['YoYGrowthSep'] = (df['Monthly Achievement(Sep)'] - df['PrevYearSep']) / df['PrevYearSep']
    df['ZoneBrand'] = df['Zone'] + '_' + df['Brand']

    feature_columns = [f'Achievement({month})' for month in months] + \
                      [f'Target({month})' for month in months] + \
                      ['PrevYearSep', 'PrevYearOct', 'YoYGrowthSep', 'ZoneBrand']
    target_column = 'Month Tgt (Oct)'

    return df, feature_columns, target_column

def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)

    xgb_model = xgb.XGBRegressor(n_estimators=100, learning_rate=0.1, random_state=42)
    xgb_model.fit(X_train_scaled, y_train)

    rf_model = RandomForestRegressor(n_estimators=100, random_state=42)
    rf_model.fit(X_train_scaled, y_train)

    return xgb_model, rf_model, scaler

def predict_sales(df, region, brand, xgb_model, rf_model, scaler, feature_columns):
    region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)].copy()

    if len(region_data) > 0:
        X_pred = region_data[feature_columns].iloc[-1:] 
        X_pred_scaled = scaler.transform(X_pred)

        xgb_pred = xgb_model.predict(X_pred_scaled)[0]
        rf_pred = rf_model.predict(X_pred_scaled)[0]
        ensemble_pred = (xgb_pred + rf_pred) / 2

        confidence_interval = 1.96 * np.std([xgb_pred, rf_pred])

        return ensemble_pred, ensemble_pred - confidence_interval, ensemble_pred + confidence_interval
    else:
        return None, None, None

# Function to generate combined report
def generate_combined_report(df, regions, brands, xgb_model, rf_model, scaler, feature_columns):
    main_table_data = [['Region', 'Brand', 'Month Target\n(Oct)', 'Monthly Achievement\n(Sep)', 'Predicted\nAchievement(Oct)', 'CI', 'RMSE']]
    
    with ThreadPoolExecutor() as executor:
        futures = []
        for region in regions:
            for brand in brands:
                futures.append(executor.submit(predict_sales, df, region, brand, xgb_model, rf_model, scaler, feature_columns))
        
        valid_data = False
        for future, (region, brand) in zip(futures, [(r, b) for r in regions for b in brands]):
            try:
                oct_achievement, lower_achievement, upper_achievement = future.result()
                if oct_achievement is not None:
                    region_data = df[(df['Zone'] == region) & (df['Brand'] == brand)]
                    if not region_data.empty:
                        oct_target = region_data['Month Tgt (Oct)'].iloc[-1]
                        sept_achievement = region_data['Monthly Achievement(Sep)'].iloc[-1]
                        
                        # Calculate RMSE using September data as a proxy
                        rmse = np.sqrt(mean_squared_error(region_data['Monthly Achievement(Sep)'], region_data['Month Tgt (Sep)']))
                        
                        main_table_data.append([
                            region, brand, f"{oct_target:.0f}", f"{sept_achievement:.0f}",
                            f"{oct_achievement:.0f}", f"({lower_achievement:.2f},\n{upper_achievement:.2f})", f"{rmse:.4f}"
                        ])
                        
                        valid_data = True
                    else:
                        st.warning(f"No data available for {region} and {brand}")
            except Exception as e:
                st.warning(f"Error processing {region} and {brand}: {str(e)}")
    
    if valid_data:
        fig, ax = plt.subplots(figsize=(12, len(main_table_data) * 0.5))
        ax.axis('off')
        table = ax.table(cellText=main_table_data[1:], colLabels=main_table_data[0], cellLoc='center', loc='center')
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
        
        plt.title("Combined Sales Predictions Report", fontsize=16, fontweight='bold', pad=20)
        plt.tight_layout()
        
        pdf_buffer = BytesIO()
        plt.savefig(pdf_buffer, format='pdf', bbox_inches='tight')
        plt.close(fig)
        
        pdf_buffer.seek(0)
        return base64.b64encode(pdf_buffer.getvalue()).decode()
    else:
        st.warning("No valid data available for any region and brand combination.")
        return None

# Streamlit app
def combined_report_app():
    st.title("📊 Combined Sales Prediction Report Generator")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    
    if uploaded_file is not None:
        with st.spinner("Loading and processing data..."):
            df = pd.read_excel(uploaded_file)
            regions = df['Zone'].unique().tolist()
            brands = df['Brand'].unique().tolist()
            
            df, feature_columns, target_column = preprocess_data(df)
            X = df[feature_columns]
            y = df[target_column]
            
            xgb_model, rf_model, scaler = train_model(X, y)
        
        st.success("Data processed and model trained successfully!")
        
        if st.button("Generate Combined Report"):
            with st.spinner("Generating combined report..."):
                combined_report_data = generate_combined_report(df, regions, brands, xgb_model, rf_model, scaler, feature_columns)
            
            if combined_report_data:
                st.success("Combined report generated successfully!")
                st.download_button(
                    label="Download Combined PDF Report",
                    data=base64.b64decode(combined_report_data),
                    file_name="combined_prediction_report.pdf",
                    mime="application/pdf"
                )
            else:
                st.error("Unable to generate combined report. Please check the warnings above for more details.")

if __name__ == "__main__":
    combined_report_app()
