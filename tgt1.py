import streamlit as st
import io
import base64
import pandas as pd
import numpy as np
from sklearn.model_selection import train_test_split
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_squared_error, r2_score
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Frame, Indenter
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Sales Prediction Dashboard", layout="wide")

# Custom CSS to improve the app's aesthetics
st.markdown("""
    <style>
    .reportview-container {
        background: #f0f2f6
    }
    .big-font {
        font-size:30px !important;
        font-weight: bold;
        color: #1E3A8A;
    }
    .stProgress > div > div > div > div {
        background-color: #1E3A8A;
    }
    </style>
    """, unsafe_allow_html=True)

@st.cache_data
def load_data(file):
    data = pd.read_excel(file)
    return data

@st.cache_resource
def train_model(X, y):
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = RandomForestRegressor(n_estimators=100, random_state=42)
    model.fit(X_train, y_train)
    return model, X_test, y_test

def create_pdf(data):
    # ... [The create_pdf function remains unchanged] ...

def main():
    st.markdown('<p class="big-font">Sales Prediction Dashboard</p>', unsafe_allow_html=True)

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        data = load_data(uploaded_file)

        features = ['Month Tgt (Oct)', 'Monthly Achievement(Sep)', 'Total Sep 2023', 'Total Oct 2023',
                    'Monthly Achievement(Apr)', 'Monthly Achievement(May)', 'Monthly Achievement(June)',
                    'Monthly Achievement(July)', 'Monthly Achievement(Aug)']

        X = data[features]
        y = data['Total Oct 2023']

        model, X_test, y_test = train_model(X, y)

        st.sidebar.header("Filters")
        selected_brands = st.sidebar.multiselect("Select Brands", data['Brand'].unique(), default=data['Brand'].unique())
        selected_zones = st.sidebar.multiselect("Select Zones", data['Zone'].unique(), default=data['Zone'].unique())

        filtered_data = data[(data['Brand'].isin(selected_brands)) & (data['Zone'].isin(selected_zones))]

        col1, col2 = st.columns(2)

        with col1:
            st.subheader("Model Performance")
            y_pred = model.predict(X_test)
            mse = mean_squared_error(y_test, y_pred)
            r2 = r2_score(y_test, y_pred)

            st.metric("Mean Squared Error", f"{mse:.2f}")
            st.metric("R-squared Score", f"{r2:.2f}")

            feature_importance = pd.DataFrame({
                'feature': features,
                'importance': model.feature_importances_
            }).sort_values('importance', ascending=False)

            fig_importance = px.bar(feature_importance, x='importance', y='feature', orientation='h',
                                    title='Feature Importance', labels={'importance': 'Importance', 'feature': 'Feature'})
            st.plotly_chart(fig_importance, use_container_width=True)

        with col2:
            st.subheader("Sales Predictions")
            X_2024 = filtered_data[features].copy()
            X_2024['Total Oct 2023'] = filtered_data['Total Oct 2023']
            predictions_2024 = model.predict(X_2024)
            filtered_data['Predicted Oct 2024'] = predictions_2024
            filtered_data['YoY Growth'] = (filtered_data['Predicted Oct 2024'] - filtered_data['Total Oct 2023']) / filtered_data['Total Oct 2023'] * 100

            fig_predictions = go.Figure()
            fig_predictions.add_trace(go.Bar(x=filtered_data['Zone'], y=filtered_data['Total Oct 2023'], name='Oct 2023 Sales'))
            fig_predictions.add_trace(go.Bar(x=filtered_data['Zone'], y=filtered_data['Predicted Oct 2024'], name='Predicted Oct 2024 Sales'))
            fig_predictions.update_layout(title='Sales Comparison: Oct 2023 vs Predicted Oct 2024', barmode='group')
            st.plotly_chart(fig_predictions, use_container_width=True)

        st.subheader("Detailed Predictions")
        st.dataframe(filtered_data[['Zone', 'Brand', 'Month Tgt (Oct)', 'Predicted Oct 2024', 'Total Oct 2023', 'YoY Growth']])

        pdf_buffer = create_pdf(filtered_data)
        pdf_data = base64.b64encode(pdf_buffer.getvalue()).decode()

        st.download_button(
            label="Download PDF Report",
            data=pdf_buffer,
            file_name="sales_predictions_oct_2024.pdf",
            mime="application/pdf"
        )
    else:
        st.info("Please upload an Excel file to start the analysis.")

if __name__ == "__main__":
    main()
