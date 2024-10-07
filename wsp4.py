import os
from datetime import datetime
from streamlit_option_menu import option_menu
import shutil
import streamlit as st
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO
import base64
import matplotlib.backends.backend_pdf
from scipy import stats
from statsmodels.tsa.arima.model import ARIMA
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet
from streamlit_lottie import st_lottie
import json
import requests
from openpyxl.utils import get_column_letter
import plotly.express as px
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak
from sklearn.model_selection import train_test_split
from sklearn.metrics import mean_squared_error
import math
import seaborn as sns
import xgboost as xgb
from io import BytesIO
import plotly.graph_objs as go
import time
from collections import OrderedDict
import re
import plotly.graph_objects as go
import plotly.express as px
from concurrent.futures import ThreadPoolExecutor
def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
import streamlit as st
import pandas as pd
import openpyxl
from collections import OrderedDict
import base64
from io import BytesIO
import numpy as np
import plotly.express as px
from statsmodels.tsa.arima.model import ARIMA
from scipy import stats
import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
from collections import OrderedDict
import plotly.express as px
import plotly.graph_objects as go
from statsmodels.tsa.arima.model import ARIMA
from scipy import stats
import base64
from io import BytesIO
import statsmodels.api as sm
from statsmodels.stats.diagnostic import het_breuschpagan, acorr_ljungbox
from statsmodels.stats.stattools import durbin_watson
from statsmodels.stats.outliers_influence import variance_inflation_factor
from statsmodels.tsa.stattools import adfuller
from statsmodels.graphics.tsaplots import plot_acf, plot_pacf
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler, PolynomialFeatures
from sklearn.linear_model import LinearRegression, Ridge, Lasso
from sklearn.tree import DecisionTreeRegressor
from sklearn.ensemble import RandomForestRegressor
from sklearn.svm import SVR
from sklearn.cluster import KMeans
from sklearn.decomposition import PCA
from sklearn.metrics import mean_squared_error, r2_score
import seaborn as sns
import matplotlib.pyplot as plt
from scipy.stats import jarque_bera, kurtosis, skew
from statsmodels.stats.stattools import omni_normtest

def excel_editor_and_analyzer():
    
    st.title("Advanced Excel Editor and Analyzer")
    tab1, tab2 = st.tabs(["Excel Editor", "Data Analyzer"])
    
    with tab1:
        excel_editor()
    
    with tab2:
        data_analyzer()

def excel_editor():
    st.header("Excel Editor")
    def create_excel_structure_html(sheet, max_rows=5):
        html = "<table class='excel-table'>"
        merged_cells = sheet.merged_cells.ranges

        for idx, row in enumerate(sheet.iter_rows(max_row=max_rows)):
            html += "<tr>"
            for cell in row:
                merged = False
                for merged_range in merged_cells:
                    if cell.coordinate in merged_range:
                        if cell.coordinate == merged_range.start_cell.coordinate:
                            rowspan = min(merged_range.max_row - merged_range.min_row + 1, max_rows - idx)
                            colspan = merged_range.max_col - merged_range.min_col + 1
                            html += f"<td rowspan='{rowspan}' colspan='{colspan}'>{cell.value}</td>"
                        merged = True
                        break
                if not merged:
                    html += f"<td>{cell.value}</td>"
            html += "</tr>"
        html += "</table>"
        return html

    # Function to get merged column groups
    def get_merged_column_groups(sheet):
        merged_groups = {}
        for merged_range in sheet.merged_cells.ranges:
            if merged_range.min_row == 1:  # Only consider merged cells in the first row (header)
                main_col = sheet.cell(1, merged_range.min_col).value
                merged_groups[main_col] = list(range(merged_range.min_col, merged_range.max_col + 1))
        return merged_groups

    # File uploader
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

    if uploaded_file is not None:
        # Read Excel file
        excel_file = openpyxl.load_workbook(uploaded_file)
        sheet = excel_file.active

        # Display original Excel structure (first 5 rows)
        st.subheader("Original Excel Structure (First 5 Rows)")
        excel_html = create_excel_structure_html(sheet, max_rows=5)
        st.markdown(excel_html, unsafe_allow_html=True)

        # Get merged column groups
        merged_groups = get_merged_column_groups(sheet)

        # Create a list of column headers, considering merged cells
        column_headers = []
        column_indices = OrderedDict()  # To store the column indices for each header
        for col in range(1, sheet.max_column + 1):
            cell_value = sheet.cell(1, col).value
            if cell_value is not None:
                column_headers.append(cell_value)
                if cell_value not in column_indices:
                    column_indices[cell_value] = []
                column_indices[cell_value].append(col - 1)  # pandas uses 0-based index
            else:
                # If the cell is empty, it's part of a merged cell, so use the previous header
                prev_header = column_headers[-1]
                column_headers.append(prev_header)
                column_indices[prev_header].append(col - 1)

        # Read as pandas DataFrame using the correct column headers
        df = pd.read_excel(uploaded_file, header=None, names=column_headers)
        df = df.iloc[1:]  # Remove the first row as it's now our header

        # Column selection for deletion
        st.subheader("Select columns to delete")
        all_columns = list(column_indices.keys())  # Use OrderedDict keys to maintain order
        cols_to_delete = st.multiselect("Choose columns to remove", all_columns)
        
        if cols_to_delete:
            columns_to_remove = []
            for col in cols_to_delete:
                columns_to_remove.extend(column_indices[col])
            
            df = df.drop(df.columns[columns_to_remove], axis=1)
            st.success(f"Deleted columns: {', '.join(cols_to_delete)}")

        # Row deletion
        st.subheader("Delete rows")
        num_rows = st.number_input("Enter the number of rows to delete from the start", min_value=0, max_value=len(df)-1, value=0)
        
        if num_rows > 0:
            df = df.iloc[num_rows:]
            st.success(f"Deleted first {num_rows} rows")
        
        # Display editable dataframe
        st.subheader("Edit Data")
        st.write("You can edit individual cell values directly in the table below:")
        
        # Replace NaN values with None and convert dataframe to a dictionary
        df_dict = df.where(pd.notnull(df), None).to_dict('records')
        
        # Use st.data_editor with the processed dictionary
        edited_data = st.data_editor(df_dict)
        
        # Convert edited data back to dataframe
        edited_df = pd.DataFrame(edited_data)
        st.subheader("Edited Data")
        st.dataframe(edited_df)
        
        # Download button
        def get_excel_download_link(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
            excel_data = output.getvalue()
            b64 = base64.b64encode(excel_data).decode()
            return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="edited_file.xlsx">Download Edited Excel File</a>'
        
        st.markdown(get_excel_download_link(edited_df), unsafe_allow_html=True)

        # New button to upload edited file to Home
        if st.button("Upload Edited File to Home"):
            # Save the edited DataFrame to session state
            st.session_state.edited_df = edited_df
            st.session_state.edited_file_name = "edited_" + uploaded_file.name
            st.success("Edited file has been uploaded to Home. Please switch to the Home tab to see the uploaded file.")

    else:
        st.info("Please upload an Excel file to begin editing.")

def data_analyzer():
    st.header("Advanced Data Analyzer")
    
    uploaded_file = st.file_uploader("Choose an Excel file for analysis", type="xlsx", key="analyser")
    
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        
        st.write("Dataset Information:")
        st.write(f"Number of rows: {df.shape[0]}")
        st.write(f"Number of columns: {df.shape[1]}")
        
        numeric_columns = df.select_dtypes(include=['float64', 'int64']).columns
        categorical_columns = df.select_dtypes(include=['object']).columns
        
        analysis_type = st.selectbox("Select analysis type", ["Univariate Analysis", "Bivariate Analysis", "Regression Analysis", "Machine Learning Models", "Advanced Statistics"])
        
        if analysis_type == "Univariate Analysis":
            univariate_analysis(df, numeric_columns, categorical_columns)
        elif analysis_type == "Bivariate Analysis":
            bivariate_analysis(df, numeric_columns)
        elif analysis_type == "Regression Analysis":
            regression_analysis(df, numeric_columns, categorical_columns)
        elif analysis_type == "Machine Learning Models":
            machine_learning_models(df, numeric_columns, categorical_columns)
        elif analysis_type == "Advanced Statistics":
            advanced_statistics(df, numeric_columns)

def univariate_analysis(df, numeric_columns, categorical_columns):
    st.subheader("Univariate Analysis")
    
    column = st.selectbox("Select a column for analysis", numeric_columns.tolist() + categorical_columns.tolist())
    
    if column in numeric_columns:
        st.write(df[column].describe())
        
        col1, col2 = st.columns(2)
        
        with col1:
            fig = go.Figure()
            fig.add_trace(go.Histogram(x=df[column], name="Histogram"))
            fig.update_layout(title=f"Histogram for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = go.Figure()
            fig.add_trace(go.Box(y=df[column], name="Box Plot"))
            fig.update_layout(title=f"Box Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        col3, col4 = st.columns(2)
        
        with col3:
            fig = go.Figure()
            fig.add_trace(go.Violin(y=df[column], box_visible=True, line_color='black', meanline_visible=True, fillcolor='lightseagreen', opacity=0.6, x0=column))
            fig.update_layout(title=f"Violin Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col4:
            fig = px.line(df, y=column, title=f"Line Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        # Additional statistics
        st.subheader("Additional Statistics")
        col5, col6, col7 = st.columns(3)
        with col5:
            st.metric("Skewness", f"{skew(df[column]):.4f}")
        with col6:
            st.metric("Kurtosis", f"{kurtosis(df[column]):.4f}")
        with col7:
            st.metric("Coefficient of Variation", f"{df[column].std() / df[column].mean():.4f}")
        
    else:
        st.write(df[column].value_counts())
        col1, col2 = st.columns(2)
        
        with col1:
            fig = px.bar(df[column].value_counts(), title=f"Bar Plot for {column}")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            fig = px.pie(df, names=column, title=f"Pie Chart for {column}")
            st.plotly_chart(fig, use_container_width=True)

def bivariate_analysis(df, numeric_columns):
    st.subheader("Bivariate Analysis")
    
    x_col = st.selectbox("Select X-axis variable", numeric_columns)
    y_col = st.selectbox("Select Y-axis variable", numeric_columns)
    
    chart_type = st.selectbox("Select chart type", ["Scatter", "Line", "Bar", "Box", "Violin", "3D Scatter", "Heatmap"])
    
    if chart_type == "Scatter":
        fig = px.scatter(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Line":
        fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Bar":
        fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Box":
        fig = px.box(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "Violin":
        fig = px.violin(df, x=x_col, y=y_col, title=f"{y_col} vs {x_col}")
    elif chart_type == "3D Scatter":
        z_col = st.selectbox("Select Z-axis variable", numeric_columns)
        fig = px.scatter_3d(df, x=x_col, y=y_col, z=z_col, title=f"3D Scatter Plot")
    elif chart_type == "Heatmap":
        corr_matrix = df[numeric_columns].corr()
        fig = px.imshow(corr_matrix, title="Correlation Heatmap")
    
    st.plotly_chart(fig, use_container_width=True)
    
    correlation = df[[x_col, y_col]].corr().iloc[0, 1]
    st.write(f"Correlation between {x_col} and {y_col}: {correlation:.4f}")
    
    # Add correlation interpretation
    st.subheader("Correlation Interpretation")
    st.write("""
    The correlation coefficient ranges from -1 to 1:
    - 1: Perfect positive correlation
    - 0: No correlation
    - -1: Perfect negative correlation
    
    Interpretation:
    - 0.00 to 0.19: Very weak correlation
    - 0.20 to 0.39: Weak correlation
    - 0.40 to 0.59: Moderate correlation
    - 0.60 to 0.79: Strong correlation
    - 0.80 to 1.00: Very strong correlation
    """)
    
    # Add correlation formula
    st.latex(r'''
    r = \frac{\sum_{i=1}^{n} (x_i - \bar{x})(y_i - \bar{y})}{\sqrt{\sum_{i=1}^{n} (x_i - \bar{x})^2} \sqrt{\sum_{i=1}^{n} (y_i - \bar{y})^2}}
    ''')
    st.write("Where:")
    st.write("- r is the correlation coefficient")
    st.write("- x_i and y_i are individual sample points")
    st.write("- x̄ and ȳ are the sample means")

def regression_analysis(df, numeric_columns, categorical_columns):
    st.subheader("Regression Analysis")
    
    regression_type = st.selectbox("Select regression type", ["Simple Linear", "Multiple Linear", "Polynomial", "Ridge", "Lasso"])
    
    y_col = st.selectbox("Select dependent variable", numeric_columns)
    x_cols = st.multiselect("Select independent variables", numeric_columns.tolist() + categorical_columns.tolist())
    
    if len(x_cols) == 0:
        st.warning("Please select at least one independent variable.")
        return
    
    X = df[x_cols]
    y = df[y_col]
    
    # Handle categorical variables
    X = pd.get_dummies(X, drop_first=True)
    
    if regression_type == "Polynomial":
        degree = st.slider("Select polynomial degree", 1, 5, 2)
        poly = PolynomialFeatures(degree=degree)
        X = poly.fit_transform(X)
    
    X = sm.add_constant(X)
    
    try:
        if regression_type == "Ridge":
            alpha = st.slider("Select alpha for Ridge regression", 0.0, 10.0, 1.0)
            model = sm.OLS(y, X).fit_regularized(alpha=alpha, L1_wt=0)
        elif regression_type == "Lasso":
            alpha = st.slider("Select alpha for Lasso regression", 0.0, 10.0, 1.0)
            model = sm.OLS(y, X).fit_regularized(alpha=alpha, L1_wt=1)
        else:
            model = sm.OLS(y, X).fit()
        
        st.write(model.summary())
        
        # Plot actual vs predicted values
        fig = px.scatter(x=y, y=model.predict(X), labels={'x': 'Actual', 'y': 'Predicted'}, title="Actual vs Predicted Values")
        fig.add_trace(go.Scatter(x=[y.min(), y.max()], y=[y.min(), y.max()], mode='lines', name='y=x'))
        st.plotly_chart(fig, use_container_width=True)
        
        # Residual plot
        residuals = model.resid
        fig = px.scatter(x=model.predict(X), y=residuals, labels={'x': 'Predicted', 'y': 'Residuals'}, title="Residual Plot")
        fig.add_hline(y=0, line_dash="dash", line_color="red")
        st.plotly_chart(fig, use_container_width=True)
        
        # Statistical tests
        st.subheader("Statistical Tests")
        
        # Normality test (Jarque-Bera)
        jb_statistic, jb_p_value = jarque_bera(residuals)
        st.write(f"Jarque-Bera Test for Normality: statistic = {jb_statistic:.4f}, p-value = {jb_p_value:.4f}")
        st.write(f"{'Reject' if jb_p_value < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
        
        # Heteroscedasticity test (Breusch-Pagan)
        _, bp_p_value, _, _ = het_breuschpagan(residuals, model.model.exog)
        st.write(f"Breusch-Pagan Test for Heteroscedasticity: p-value = {bp_p_value:.4f}")
        st.write(f"{'Reject' if bp_p_value < 0.05 else 'Fail to reject'} the null hypothesis of homoscedasticity at 5% significance level.")
        dw_statistic = durbin_watson(residuals)
        st.write(f"Durbin-Watson Test for Autocorrelation: {dw_statistic:.4f}")
        st.write("Values close to 2 suggest no autocorrelation, while values toward 0 or 4 suggest positive or negative autocorrelation.")
        
        # Multicollinearity (VIF)
        vif_data = pd.DataFrame()
        vif_data["Variable"] = X.columns
        vif_data["VIF"] = [variance_inflation_factor(X.values, i) for i in range(X.shape[1])]
        st.write("Variance Inflation Factors (VIF) for Multicollinearity:")
        st.write(vif_data)
        st.write("VIF > 5 suggests high multicollinearity.")
        
        # Add regression formulas and explanations
        st.subheader("Regression Formulas")
        if regression_type == "Simple Linear":
            st.latex(r'y = \beta_0 + \beta_1x + \epsilon')
        elif regression_type == "Multiple Linear":
            st.latex(r'y = \beta_0 + \beta_1x_1 + \beta_2x_2 + ... + \beta_nx_n + \epsilon')
        elif regression_type == "Polynomial":
            st.latex(r'y = \beta_0 + \beta_1x + \beta_2x^2 + ... + \beta_nx^n + \epsilon')
        elif regression_type == "Ridge":
            st.latex(r'\min_{\beta} \sum_{i=1}^n (y_i - \beta_0 - \sum_{j=1}^p \beta_jx_{ij})^2 + \lambda \sum_{j=1}^p \beta_j^2')
        elif regression_type == "Lasso":
            st.latex(r'\min_{\beta} \sum_{i=1}^n (y_i - \beta_0 - \sum_{j=1}^p \beta_jx_{ij})^2 + \lambda \sum_{j=1}^p |\beta_j|')
        
        st.write("Where:")
        st.write("- y is the dependent variable")
        st.write("- x, x_1, x_2, ..., x_n are independent variables")
        st.write("- β_0, β_1, β_2, ..., β_n are regression coefficients")
        st.write("- ε is the error term")
        st.write("- λ is the regularization parameter (for Ridge and Lasso)")
        
    except Exception as e:
        st.error(f"An error occurred during regression analysis: {str(e)}")
        st.write("This error might be due to multicollinearity, insufficient data, or other issues in the dataset.")
        st.write("Try selecting different variables or using a different regression type.")

def machine_learning_models(df, numeric_columns, categorical_columns):
    st.subheader("Machine Learning Models")
    
    model_type = st.selectbox("Select model type", ["Supervised", "Unsupervised"])
    
    if model_type == "Supervised":
        supervised_models(df, numeric_columns, categorical_columns)
    else:
        unsupervised_models(df, numeric_columns)

def supervised_models(df, numeric_columns, categorical_columns):
    st.write("Supervised Learning Models")
    
    y_col = st.selectbox("Select target variable", numeric_columns)
    x_cols = st.multiselect("Select features", numeric_columns.tolist() + categorical_columns.tolist())
    
    if len(x_cols) == 0:
        st.warning("Please select at least one feature.")
        return
    
    X = df[x_cols]
    y = df[y_col]
    
    # Handle categorical variables
    X = pd.get_dummies(X, drop_first=True)
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)
    
    models = {
        "Linear Regression": LinearRegression(),
        "Decision Tree": DecisionTreeRegressor(),
        "Random Forest": RandomForestRegressor(),
        "SVR": SVR()
    }
    
    selected_model = st.selectbox("Select a model", list(models.keys()))
    
    try:
        model = models[selected_model]
        model.fit(X_train_scaled, y_train)
        
        y_pred = model.predict(X_test_scaled)
        
        mse = mean_squared_error(y_test, y_pred)
        r2 = r2_score(y_test, y_pred)
        
        st.write(f"Mean Squared Error: {mse:.4f}")
        st.write(f"R-squared Score: {r2:.4f}")
        
        fig = px.scatter(x=y_test, y=y_pred, labels={'x': 'Actual', 'y': 'Predicted'}, title="Actual vs Predicted Values")
        fig.add_trace(go.Scatter(x=[y_test.min(), y_test.max()], y=[y_test.min(), y_test.max()], mode='lines', name='y=x'))
        st.plotly_chart(fig, use_container_width=True)
        
        # Feature importance (for tree-based models)
        if selected_model in ["Decision Tree", "Random Forest"]:
            feature_importance = pd.DataFrame({
                'feature': X.columns,
                'importance': model.feature_importances_
            }).sort_values('importance', ascending=False)
            
            st.write("Feature Importance:")
            fig = px.bar(feature_importance, x='feature', y='importance', title="Feature Importance")
            st.plotly_chart(fig, use_container_width=True)
        
        # Add model formulas and explanations
        st.subheader("Model Formulas and Explanations")
        if selected_model == "Linear Regression":
            st.latex(r'y = \beta_0 + \beta_1x_1 + \beta_2x_2 + ... + \beta_nx_n')
            st.write("Linear Regression finds the best-fitting linear relationship between the target variable and the features.")
        elif selected_model == "Decision Tree":
            st.write("Decision Trees make predictions by learning decision rules inferred from the data features.")
            st.image("https://scikit-learn.org/stable/_images/iris_dtc.png", caption="Example of a Decision Tree")
        elif selected_model == "Random Forest":
            st.write("Random Forest is an ensemble of Decision Trees, where each tree is trained on a random subset of the data and features.")
            st.image("https://scikit-learn.org/stable/_images/plot_forest_importances_faces_001.png", caption="Example of Random Forest Feature Importance")
        elif selected_model == "SVR":
            st.latex(r'\min_{w, b, \xi} \frac{1}{2} \|w\|^2 + C \sum_{i=1}^n \xi_i')
            st.write("Support Vector Regression (SVR) finds a function that deviates from y by a value no greater than ε for each training point x.")
    
    except Exception as e:
        st.error(f"An error occurred during model training: {str(e)}")
        st.write("This error might be due to insufficient data, incompatible data types, or other issues in the dataset.")
        st.write("Try selecting different variables or using a different model.")

def unsupervised_models(df, numeric_columns):
    st.write("Unsupervised Learning Models")
    
    x_cols = st.multiselect("Select features for clustering", numeric_columns)
    
    if len(x_cols) == 0:
        st.warning("Please select at least one feature.")
        return
    
    X = df[x_cols]
    
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)
    
    n_clusters = st.slider("Select number of clusters", 2, 10, 3)
    
    try:
        kmeans = KMeans(n_clusters=n_clusters, random_state=42)
        cluster_labels = kmeans.fit_predict(X_scaled)
        
        df_clustered = df.copy()
        df_clustered['Cluster'] = cluster_labels
        
        if len(x_cols) >= 2:
            fig = px.scatter(df_clustered, x=x_cols[0], y=x_cols[1], color='Cluster', title="K-means Clustering")
            st.plotly_chart(fig, use_container_width=True)
        
        st.write("Cluster Centers:")
        cluster_centers = scaler.inverse_transform(kmeans.cluster_centers_)
        st.write(pd.DataFrame(cluster_centers, columns=x_cols))
        
        # Elbow method for optimal number of clusters
        inertias = []
        k_range = range(1, 11)
        for k in k_range:
            kmeans = KMeans(n_clusters=k, random_state=42)
            kmeans.fit(X_scaled)
            inertias.append(kmeans.inertia_)
        
        fig = px.line(x=k_range, y=inertias, title="Elbow Method for Optimal k",
                      labels={'x': 'Number of Clusters (k)', 'y': 'Inertia'})
        st.plotly_chart(fig, use_container_width=True)
        
        # PCA
        st.subheader("Principal Component Analysis (PCA)")
        n_components = st.slider("Select number of components", 2, min(len(x_cols), 10), 2)
        pca = PCA(n_components=n_components)
        pca_result = pca.fit_transform(X_scaled)
        
        df_pca = pd.DataFrame(data=pca_result, columns=[f'PC{i+1}' for i in range(n_components)])
        
        fig = px.scatter(df_pca, x='PC1', y='PC2', title="PCA Visualization")
        st.plotly_chart(fig, use_container_width=True)
        
        explained_variance_ratio = pca.explained_variance_ratio_
        cumulative_variance_ratio = np.cumsum(explained_variance_ratio)
        
        fig = go.Figure()
        fig.add_trace(go.Bar(x=range(1, n_components+1), y=explained_variance_ratio, name='Individual'))
        fig.add_trace(go.Scatter(x=range(1, n_components+1), y=cumulative_variance_ratio, mode='lines+markers', name='Cumulative'))
        fig.update_layout(title='Explained Variance Ratio', xaxis_title='Principal Components', yaxis_title='Explained Variance Ratio')
        st.plotly_chart(fig, use_container_width=True)
        
        st.write("Explained Variance Ratio:")
        st.write(pd.DataFrame({'PC': range(1, n_components+1), 'Explained Variance Ratio': explained_variance_ratio, 'Cumulative Variance Ratio': cumulative_variance_ratio}))
        
        # Add formulas and explanations
        st.subheader("K-means Clustering Formula")
        st.latex(r'\min_{S} \sum_{i=1}^{k} \sum_{x \in S_i} \|x - \mu_i\|^2')
        st.write("Where:")
        st.write("- S is the set of clusters")
        st.write("- k is the number of clusters")
        st.write("- x is a data point")
        st.write("- μ_i is the mean of points in S_i")
        
        st.subheader("PCA Formula")
        st.latex(r'X = U\Sigma V^T')
        st.write("Where:")
        st.write("- X is the original data matrix")
        st.write("- U is the left singular vectors (eigenvectors of XX^T)")
        st.write("- Σ is a diagonal matrix of singular values")
        st.write("- V^T is the right singular vectors (eigenvectors of X^TX)")
    
    except Exception as e:
        st.error(f"An error occurred during unsupervised learning: {str(e)}")
        st.write("This error might be due to insufficient data, incompatible data types, or other issues in the dataset.")
        st.write("Try selecting different variables or adjusting the number of clusters/components.")

def advanced_statistics(df, numeric_columns):
    st.subheader("Advanced Statistics")
    
    column = st.selectbox("Select a column for advanced statistics", numeric_columns)
    
    st.write("Descriptive Statistics:")
    st.write(df[column].describe())
    
    st.subheader("Normality Tests")
    
    # Shapiro-Wilk Test
    shapiro_stat, shapiro_p = stats.shapiro(df[column])
    st.write(f"Shapiro-Wilk Test: statistic = {shapiro_stat:.4f}, p-value = {shapiro_p:.4f}")
    st.write(f"{'Reject' if shapiro_p < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
    
    # Anderson-Darling Test
    anderson_result = stats.anderson(df[column])
    st.write("Anderson-Darling Test:")
    st.write(f"Statistic: {anderson_result.statistic:.4f}")
    for i in range(len(anderson_result.critical_values)):
        sl, cv = anderson_result.significance_level[i], anderson_result.critical_values[i]
        st.write(f"At {sl}% significance level: critical value = {cv:.4f}")
        if anderson_result.statistic < cv:
            st.write(f"The null hypothesis of normality is not rejected at {sl}% significance level.")
        else:
            st.write(f"The null hypothesis of normality is rejected at {sl}% significance level.")
    
    # Jarque-Bera Test
    jb_stat, jb_p = stats.jarque_bera(df[column])
    st.write(f"Jarque-Bera Test: statistic = {jb_stat:.4f}, p-value = {jb_p:.4f}")
    st.write(f"{'Reject' if jb_p < 0.05 else 'Fail to reject'} the null hypothesis of normality at 5% significance level.")
    
    # Q-Q Plot
    fig, ax = plt.subplots()
    stats.probplot(df[column], dist="norm", plot=ax)
    ax.set_title("Q-Q Plot")
    st.pyplot(fig)
    
    st.subheader("Time Series Analysis")
    
    # Augmented Dickey-Fuller Test for Stationarity
    adf_result = adfuller(df[column])
    st.write("Augmented Dickey-Fuller Test:")
    st.write(f"ADF Statistic: {adf_result[0]:.4f}")
    st.write(f"p-value: {adf_result[1]:.4f}")
    for key, value in adf_result[4].items():
        st.write(f"Critical Value ({key}): {value:.4f}")
    st.write(f"{'Reject' if adf_result[1] < 0.05 else 'Fail to reject'} the null hypothesis of a unit root at 5% significance level.")
    
    # ACF and PACF plots
    fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(10, 10))
    plot_acf(df[column], ax=ax1)
    plot_pacf(df[column], ax=ax2)
    ax1.set_title("Autocorrelation Function (ACF)")
    ax2.set_title("Partial Autocorrelation Function (PACF)")
    st.pyplot(fig)
    
    st.subheader("Distribution Fitting")
    
    # Fit normal distribution
    mu, sigma = stats.norm.fit(df[column])
    x = np.linspace(df[column].min(), df[column].max(), 100)
    y = stats.norm.pdf(x, mu, sigma)
    
    fig, ax = plt.subplots()
    ax.hist(df[column], density=True, alpha=0.7, bins='auto')
    ax.plot(x, y, 'r-', lw=2, label='Normal fit')
    ax.set_title(f"Distribution Fitting for {column}")
    ax.legend()
    st.pyplot(fig)
    
    st.write(f"Fitted Normal Distribution: μ = {mu:.4f}, σ = {sigma:.4f}")
    
    # Kolmogorov-Smirnov Test
    ks_statistic, ks_p_value = stats.kstest(df[column], 'norm', args=(mu, sigma))
    st.write("Kolmogorov-Smirnov Test:")
    st.write(f"Statistic: {ks_statistic:.4f}")
    st.write(f"p-value: {ks_p_value:.4f}")
    st.write(f"{'Reject' if ks_p_value < 0.05 else 'Fail to reject'} the null hypothesis that the data comes from the fitted normal distribution at 5% significance level.")
def create_stats_pdf(stats_data, district):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph(f"Descriptive Statistics for {district}", styles['Title'])
    elements.append(title)

    data = [['Brand', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']]
    for brand, stats in stats_data.items():
        row = [brand]
        for stat in ['Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']:
            value = stats[stat]
            if isinstance(value, (int, float)):
                row.append(f"{value:.2f}")
            else:
                row.append(str(value))
        data.append(row)

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer
def create_prediction_pdf(prediction_data, district):
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    elements = []

    styles = getSampleStyleSheet()
    title = Paragraph(f"Price Predictions for {district}", styles['Title'])
    elements.append(title)

    data = [['Brand', 'Predicted Price', 'Lower CI', 'Upper CI']]
    for brand, pred in prediction_data.items():
        row = [brand, f"{pred['forecast']:.2f}", f"{pred['lower_ci']:.2f}", f"{pred['upper_ci']:.2f}"]
        data.append(row)

    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 14),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 12),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

st.set_page_config(page_title="WSP Analysis",page_icon="📊", layout="wide")

# [Keep the existing custom CSS here]
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
        color: brown;
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

# [Keep the existing transform_data, plot_district_graph, process_file, and update_week_name functions]
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
            price_diff_label = price_diff
            if np.isnan(price_diff):
               price_diff = 'NA'
            label = f"{brand} ({price_diff if isinstance(price_diff, str) else f'{price_diff:.0f}'})"
            plt.plot(week_names, brand_prices, marker='o', linestyle='-', label=label)
            for week, price in zip(week_names, brand_prices):
                if not np.isnan(price):
                    plt.text(week, price, str(round(price)), fontsize=10)
        plt.grid(False)
        plt.xlabel('Month/Week', weight='bold')
        reference_week = week_names[diff_week]
        last_week = week_names[-1]
        
        explanation_text = f"***Numbers in brackets next to brand names show the price difference between {reference_week} and {last_week}.***"
        plt.annotate(explanation_text, 
                     xy=(0, -0.23), xycoords='axes fraction', 
                     ha='left', va='center', fontsize=8, style='italic', color='deeppink',
                     bbox=dict(facecolor="#f0f8ff", edgecolor='none', alpha=0.7, pad=3))
        
        region_name = district_df['REGION'].iloc[0]
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
def update_week_name(index):
    def callback():
        if index < len(st.session_state.week_names_input):
            st.session_state.week_names_input[index] = st.session_state[f'week_{index}']
        else:
            st.warning(f"Attempted to update week {index + 1}, but only {len(st.session_state.week_names_input)} weeks are available.")
        st.session_state.all_weeks_filled = all(st.session_state.week_names_input)
    return callback


def load_lottie_url(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()
    #-webkit-background-clip: text;
        #-webkit-text-fill-color: transparent;
def Home():
    # Custom CSS with more modern and professional styling
    st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;700&display=swap');
    
    body {
        font-family: 'Roboto', sans-serif;
        background-color: #f5f7fa;
        color: #333;
    }
    .title {
        font-size: 3.5rem;
        font-weight: 700;
        color: brown;
        text-align: center;
        padding: 2rem 0;
        margin-bottom: 2rem;
        background: linear-gradient(to right, #f0f8ff, #e6f3ff);
        
    }
    .subtitle {
        font-size: 1.5rem;
        font-weight: 300;
        color: #34495e;
        text-align: center;
        margin-bottom: 2rem;
    }
    .section-box {
        background-color: #ffffff;
        border-radius: 8px;
        padding: 2rem;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 12px rgba(0, 0, 0, 0.15);
    }
    .upload-section {
        background: linear-gradient(120deg, #a1c4fd 0%, #c2e9fb 100%);
        padding: 2rem;
        border-radius: 8px;
        margin-bottom: 2rem;
    }
    .btn-primary {
        background-color: #3498db;
        color: brown;
        padding: 0.5rem 1rem;
        border-radius: 5px;
        border: none;
        cursor: pointer;
        transition: background-color 0.3s ease;
    }
    .btn-primary:hover {
        background-color: #2980b9;
    }
    </style>
    """, unsafe_allow_html=True)

    # Main title and subtitle
    st.markdown('<h1 class="title">Statistica Pro</h1>', unsafe_allow_html=True)
    st.markdown('<p class="subtitle">Analyze, Visualize, Optimize.</p>', unsafe_allow_html=True)

    # Load and display Lottie animation
    lottie_url = "https://assets9.lottiefiles.com/packages/lf20_jcikwtux.json"
    lottie_json = load_lottie_url(lottie_url)

    col1, col2 = st.columns([1, 2])
    with col1:
        st_lottie(lottie_json, height=250, key="home_animation")
    with col2:
        st.markdown("""
        <div class="section-box">
        <h3>Welcome to Your Data Analysis Journey!</h3>
        <p>Our interactive dashboard empowers you to:</p>
        <ul>
            <li>Upload and process your WSP data effortlessly</li>
            <li>Visualize trends across different brands and regions</li>
            <li>Generate descriptive statistics and predictions</li>
            <li>Make data-driven decisions with confidence</li>
        </ul>
        </div>
        """, unsafe_allow_html=True)

    # How to use section
    st.markdown("""
    <div class="section-box">
    <h3>How to Use This Dashboard</h3>
    <ol>
        <li><strong>Upload Your Data:</strong> Start by uploading your Excel file containing the WSP data.</li>
        <li><strong>Enter Week Names:</strong> Provide names for each week column in your dataset.</li>
        <li><strong>Choose Your Analysis:</strong> Navigate to either the WSP Analysis Dashboard or Descriptive Statistics and Prediction sections.</li>
        <li><strong>Customize and Explore:</strong> Select your analysis parameters and generate valuable insights!</li>
    </ol>
    </div>
    """, unsafe_allow_html=True)

    # File upload section
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    st.subheader("Upload Your Data")

    if 'file_processed' not in st.session_state:
        st.session_state.file_processed = False
    if 'file_ready' not in st.session_state:
        st.session_state.file_ready = False

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"],key="wsp_data")
    if 'edited_df' in st.session_state and 'edited_file_name' in st.session_state and not st.session_state.edited_df.empty:
        st.success(f"Edited file uploaded: {st.session_state.edited_file_name}")
        if st.button("Process Edited File", key="process_edited"):
            process_uploaded_file(st.session_state.edited_df)

    elif uploaded_file:
        st.success(f"File uploaded: {uploaded_file.name}")
        if st.button("Process Uploaded File", key="process_uploaded"):
            process_uploaded_file(uploaded_file)

    if st.session_state.file_ready:
        st.markdown("### Enter Week Names")
        num_weeks = st.session_state.num_weeks
        num_columns = min(4, num_weeks)  # Limit to 4 columns for better layout
        week_cols = st.columns(num_columns)

        for i in range(num_weeks):
            with week_cols[i % num_columns]:
                st.session_state.week_names_input[i] = st.text_input(
                    f'Week {i+1}', 
                    value=st.session_state.week_names_input[i],
                    key=f'week_{i}'
                )
        
        if st.button("Confirm Week Names", key="confirm_weeks"):
            if all(st.session_state.week_names_input):
                st.session_state.file_processed = True
                st.success("File processed successfully! You can now proceed to the analysis sections.")
            else:
                st.warning("Please fill in all week names before confirming.")

    if st.session_state.file_processed:
        st.success("File processed successfully! You can now proceed to the analysis sections.")
    else:
        st.info("Please upload a file and fill in all week names to proceed with the analysis.")

    st.markdown('</div>', unsafe_allow_html=True)

    # Help section
    st.markdown("""
    <div class="section-box">
    <h3>Need Assistance?</h3>
    <p>If you have any questions or need help using the dashboard, our support team is here for you. Don't hesitate to reach out!</p>
    <p>Email: prasoon.bajpai@lc.jkmail.com</p>
    <p>Phone: +91-9219393559</p>
    </div>
    """, unsafe_allow_html=True)

    # Footer
    st.markdown("""
    <div style="text-align: center; margin-top: 2rem; padding: 1rem; background-color: #34495e; color: #ecf0f1;">
    <p>© 2024 WSP Analysis Dashboard. All rights reserved.</p>
    </div>
    """, unsafe_allow_html=True)

def process_uploaded_file(uploaded_file):
    if (isinstance(uploaded_file, pd.DataFrame) or uploaded_file) and not st.session_state.file_processed:
        try:
            if isinstance(uploaded_file, pd.DataFrame):
                # Convert DataFrame to Excel file in memory
                buffer = BytesIO()
                uploaded_file.to_excel(buffer, index=False)
                buffer.seek(0)
                file_content = buffer.getvalue()
            else:
                file_content = uploaded_file.read()

            # Load workbook to check for hidden columns
            wb = openpyxl.load_workbook(BytesIO(file_content))
            ws = wb.active
            hidden_cols = [idx for idx, col in enumerate(ws.column_dimensions, 1) if ws.column_dimensions[col].hidden]

            # Read Excel file with header=2 for both cases
            df = pd.read_excel(BytesIO(file_content), header=2)
            df = df.dropna(axis=1, how='all')
            df = df.drop(columns=df.columns[hidden_cols], errors='ignore')

            if df.empty:
                st.error("The uploaded file resulted in an empty dataframe. Please check the file content.")
            else:
                st.session_state.df = df
                brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                brand_columns = [col for col in st.session_state.df.columns if any(brand in str(col) for brand in brands)]
                num_weeks = len(brand_columns) // len(brands)
                
                if num_weeks > 0:
                    if 'week_names_input' not in st.session_state or len(st.session_state.week_names_input) != num_weeks:
                        st.session_state.week_names_input = [''] * num_weeks
                    
                    st.session_state.num_weeks = num_weeks
                    st.session_state.file_ready = True
                else:
                    st.warning("No weeks detected in the uploaded file. Please check the file content.")
                    st.session_state.week_names_input = []
                    st.session_state.file_processed = False
        except Exception as e:
            st.error(f"Error processing file: {e}")
            st.exception(e)
            st.session_state.file_processed = False
import streamlit as st
from streamlit_option_menu import option_menu

def wsp_analysis_dashboard():
    st.markdown("""
    <style>
    .title {
        font-size: 50px;
        font-weight: bold;
        color: brown;
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
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .stSelectbox, .stMultiSelect {
        background-color: white;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>WSP Analysis Dashboard</span></div>', unsafe_allow_html=True)

    if not st.session_state.file_processed:
        st.warning("Please upload a file and fill in all week names in the Home section before using this dashboard.")
        return

    st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)
    
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Analysis Settings")
    
    st.session_state.diff_week = st.slider("Select Week for Difference Calculation", 
                                           min_value=0, 
                                           max_value=len(st.session_state.week_names_input) - 1, 
                                           value=st.session_state.diff_week, 
                                           key="diff_week_slider") 
    download_pdf = st.checkbox("Download Plots as PDF",value=True)   
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
    if selected_region in ["Rajasthan", "Madhya Pradesh(West)","Madhya Pradesh(East)","Chhattisgarh","Maharashtra(East)","Odisha","North-I","North-II","Gujarat"]:
        suggested_districts = []
        
        if selected_region == "Rajasthan":
            rajasthan_districts = ["Alwar", "Jodhpur", "Udaipur", "Jaipur", "Kota", "Bikaner"]
            suggested_districts = [d for d in rajasthan_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(West)":
            mp_west_districts = ["Indore", "Neemuch","Ratlam","Dhar"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(East)":
            mp_west_districts = ["Jabalpur","Balaghat","Chhindwara"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Chhattisgarh":
            mp_west_districts = ["Durg","Raipur","Bilaspur","Raigarh","Rajnandgaon"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Maharashtra(East)":
            mp_west_districts = ["Nagpur","Gondiya"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Odisha":
            mp_west_districts = ["Cuttack","Sambalpur","Khorda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-I":
            mp_west_districts = ["East","Gurugram","Sonipat","Hisar","Yamunanagar","Bathinda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-II":
            mp_west_districts = ["Ghaziabad","Meerut"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Gujarat":
            mp_west_districts = ["Ahmadabad","Mahesana","Rajkot","Vadodara","Surat"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        
        
        
        if suggested_districts:
            st.markdown(f"### Suggested Districts for {selected_region}")
            select_all = st.checkbox(f"Select all suggested districts for {selected_region}")
            
            if select_all:
                selected_districts = st.multiselect("Select District(s)", district_names, default=suggested_districts, key="district_select")
            else:
                selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
        else:
            selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    else:
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    

    st.markdown('</div>', unsafe_allow_html=True)

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

    else:
        st.warning("Please upload a file in the Home section before using this dashboard.")

# Make sure to import the required libraries and define the necessary functions (transform_data, plot_district_graph) elsewhere in your code.
def descriptive_statistics_and_prediction():
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
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .stSelectbox, .stMultiSelect {
        background-color: white;
        border-radius: 8px;
        margin-bottom: 10px;
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    .stats-box {
        background-color: #e6f3ff;
        border-radius: 8px;
        padding: 15px;
        margin-bottom: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>Descriptive Statistics and Prediction</span></div>', unsafe_allow_html=True)

    if not st.session_state.file_processed:
        st.warning("Please upload a file in the Home section before using this feature.")
        return

    st.session_state.df = transform_data(st.session_state.df, st.session_state.week_names_input)

    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    st.subheader("Analysis Settings")

    col1, col2 = st.columns(2)
    with col1:
        zone_names = st.session_state.df["Zone"].unique().tolist()
        selected_zone = st.selectbox("Select Zone", zone_names, key="stats_zone_select")
    with col2:
        filtered_df = st.session_state.df[st.session_state.df["Zone"] == selected_zone]
        region_names = filtered_df["REGION"].unique().tolist()
        selected_region = st.selectbox("Select Region", region_names, key="stats_region_select")
    

    filtered_df = filtered_df[filtered_df["REGION"] == selected_region]
    district_names = filtered_df["Dist Name"].unique().tolist()
    
    if selected_region in ["Rajasthan", "Madhya Pradesh(West)","Madhya Pradesh(East)","Chhattisgarh","Maharashtra(East)","Odisha","North-I","North-II","Gujarat"]:
        suggested_districts = []
        
        if selected_region == "Rajasthan":
            rajasthan_districts = ["Alwar", "Jodhpur", "Udaipur", "Jaipur", "Kota", "Bikaner"]
            suggested_districts = [d for d in rajasthan_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(West)":
            mp_west_districts = ["Indore", "Neemuch","Ratlam","Dhar"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Madhya Pradesh(East)":
            mp_west_districts = ["Jabalpur","Balaghat","Chhindwara"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Chhattisgarh":
            mp_west_districts = ["Durg","Raipur","Bilaspur","Raigarh","Rajnandgaon"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Maharashtra(East)":
            mp_west_districts = ["Nagpur","Gondiya"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Odisha":
            mp_west_districts = ["Cuttack","Sambalpur","Khorda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-I":
            mp_west_districts = ["East","Gurugram","Sonipat","Hisar","Yamunanagar","Bathinda"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "North-II":
            mp_west_districts = ["Ghaziabad","Meerut"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        elif selected_region == "Gujarat":
            mp_west_districts = ["Ahmadabad","Mahesana","Rajkot","Vadodara","Surat"]
            suggested_districts = [d for d in mp_west_districts if d in district_names]
        
        
        
        if suggested_districts:
            st.markdown(f"### Suggested Districts for {selected_region}")
            select_all = st.checkbox(f"Select all suggested districts for {selected_region}")
            
            if select_all:
                selected_districts = st.multiselect("Select District(s)", district_names, default=suggested_districts, key="district_select")
            else:
                selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
        else:
            selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    else:
        selected_districts = st.multiselect("Select District(s)", district_names, key="district_select")
    

    st.markdown('</div>', unsafe_allow_html=True)


    if selected_districts:
        # Add a button to download all stats and predictions in one PDF
        if len(selected_districts) > 1:
            if st.checkbox("Download All Stats and Predictions",value=True):
                all_stats_pdf = BytesIO()
                pdf = SimpleDocTemplate(all_stats_pdf, pagesize=letter)
                elements = []
                
                for district in selected_districts:
                    elements.append(Paragraph(f"Statistics and Predictions for {district}", getSampleStyleSheet()['Title']))
                    district_df = filtered_df[filtered_df["Dist Name"] == district]
                    
                    brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
                    stats_data = {}
                    prediction_data = {}
                    
                    for brand in brands:
                        brand_data = district_df[[col for col in district_df.columns if brand in col]].values.flatten()
                        brand_data = brand_data[~np.isnan(brand_data)]
                        
                        if len(brand_data) > 0:
                            stats_data[brand] = pd.DataFrame({
                                'Mean': [np.mean(brand_data)],
                                'Median': [np.median(brand_data)],
                                'Std Dev': [np.std(brand_data)],
                                'Min': [np.min(brand_data)],
                                'Max': [np.max(brand_data)],
                                'Skewness': [stats.skew(brand_data)],
                                'Kurtosis': [stats.kurtosis(brand_data)],
                                'Range': [np.ptp(brand_data)],
                                'IQR': [np.percentile(brand_data, 75) - np.percentile(brand_data, 25)]
                            }).iloc[0]

                            if len(brand_data) > 2:
                                model = ARIMA(brand_data, order=(1,1,1))
                                model_fit = model.fit()
                                forecast = model_fit.forecast(steps=1)
                                confidence_interval = model_fit.get_forecast(steps=1).conf_int()
                                prediction_data[brand] = {
                                    'forecast': forecast[0],
                                    'lower_ci': confidence_interval[0, 0],
                                    'upper_ci': confidence_interval[0, 1]
                                }
                    
                    elements.append(Paragraph("Descriptive Statistics", getSampleStyleSheet()['Heading2']))
                    elements.append(create_stats_table(stats_data))
                    elements.append(Paragraph("Price Predictions", getSampleStyleSheet()['Heading2']))
                    elements.append(create_prediction_table(prediction_data))
                    elements.append(PageBreak())
                
                pdf.build(elements)
                st.download_button(
                    label="Download All Stats and Predictions PDF",
                    data=all_stats_pdf.getvalue(),
                    file_name=f"{selected_districts}stats_and_predictions.pdf",
                    mime="application/pdf"
                )

        st.markdown('<div class="section-box">', unsafe_allow_html=True)
        st.markdown("### Descriptive Statistics")
        
        for district in selected_districts:
            st.subheader(f"{district}")
            district_df = filtered_df[filtered_df["Dist Name"] == district]
            
            brands = ['UTCL', 'JKS', 'JKLC', 'Ambuja', 'Wonder', 'Shree']
            stats_data = {}
            prediction_data = {}
            
            for brand in brands:
                st.markdown(f'<div class="stats-box">', unsafe_allow_html=True)
                st.markdown(f"#### {brand}")
                brand_data = district_df[[col for col in district_df.columns if brand in col]].values.flatten()
                brand_data = brand_data[~np.isnan(brand_data)]
                
                if len(brand_data) > 0:
                    basic_stats = pd.DataFrame({
                        'Mean': [np.mean(brand_data)],
                        'Median': [np.median(brand_data)],
                        'Std Dev': [np.std(brand_data)],
                        'Min': [np.min(brand_data)],
                        'Max': [np.max(brand_data)],
                        'Skewness': [stats.skew(brand_data)],
                        'Kurtosis': [stats.kurtosis(brand_data)],
                        'Range': [np.ptp(brand_data)],
                        'IQR': [np.percentile(brand_data, 75) - np.percentile(brand_data, 25)]
                    })
                    st.dataframe(basic_stats)
                    stats_data[brand] = basic_stats.iloc[0]

                    # ARIMA prediction for next week
                    if len(brand_data) > 2:  # Need at least 3 data points for ARIMA
                        model = ARIMA(brand_data, order=(1,1,1))
                        model_fit = model.fit()
                        forecast = model_fit.forecast(steps=1)
                        confidence_interval = model_fit.get_forecast(steps=1).conf_int()
                        st.markdown(f"Predicted price for next week: {forecast[0]:.2f}")
                        st.markdown(f"95% Confidence Interval: [{confidence_interval[0, 0]:.2f}, {confidence_interval[0, 1]:.2f}]")
                        prediction_data[brand] = {
                            'forecast': forecast[0],
                            'lower_ci': confidence_interval[0, 0],
                            'upper_ci': confidence_interval[0, 1]
                        }
                else:
                    st.warning(f"No data available for {brand} in this district.")
                st.markdown('</div>', unsafe_allow_html=True)

            # Create download buttons for stats and predictions
            stats_pdf = create_stats_pdf(stats_data, district)
            predictions_pdf = create_prediction_pdf(prediction_data, district)

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    label="Download Statistics PDF",
                    data=stats_pdf,
                    file_name=f"{district}_statistics.pdf",
                    mime="application/pdf"
                )
            with col2:
                st.download_button(
                    label="Download Predictions PDF",
                    data=predictions_pdf,
                    file_name=f"{district}_predictions.pdf",
                    mime="application/pdf"
                )
        st.markdown('</div>', unsafe_allow_html=True)

def create_stats_table(stats_data):
    data = [['Brand', 'Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']]
    for brand, stats in stats_data.items():
        row = [brand]
        for stat in ['Mean', 'Median', 'Std Dev', 'Min', 'Max', 'Skewness', 'Kurtosis', 'Range', 'IQR']:
            value = stats[stat]
            if isinstance(value, (int, float)):
                row.append(f"{value:.2f}")
            else:
                row.append(str(value))
        data.append(row)
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    return table

def create_prediction_table(prediction_data):
    data = [['Brand', 'Predicted Price', 'Lower CI', 'Upper CI']]
    for brand, pred in prediction_data.items():
        row = [brand, f"{pred['forecast']:.2f}", f"{pred['lower_ci']:.2f}", f"{pred['upper_ci']:.2f}"]
        data.append(row)
    
    table = Table(data)
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 10),
        ('TOPPADDING', (0, 1), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 1), (-1, -1), 6),
        ('GRID', (0, 0), (-1, -1), 1, colors.black)
    ]))
    return table
from urllib.parse import quote
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
def xgboost_explanation():
    st.title("Understanding XGBoost")
    st.write("""
    XGBoost (eXtreme Gradient Boosting) is an advanced implementation of gradient boosting algorithms. 
    It's known for its speed and performance, particularly with structured/tabular data.
    """)

    st.header("Key Concepts")
    st.subheader("1. Ensemble Learning")
    st.write("""
    XGBoost is an ensemble learning method. It combines multiple weak learners (typically decision trees) 
    to create a strong predictor.
    """)

    st.subheader("2. Gradient Boosting")
    st.write("""
    XGBoost builds trees sequentially, with each new tree correcting the errors of the combined existing trees.
    """)

    st.latex(r'''
    F_m(x) = F_{m-1}(x) + \gamma_m h_m(x)
    ''')
    st.write("""
    Where:
    - F_m(x) is the model after m iterations
    - h_m(x) is the new tree
    - γ_m is the weight of the new tree
    """)

    st.subheader("3. Loss Function and Gradient")
    st.write("""
    XGBoost aims to minimize a loss function. The gradient of the loss function is used to identify the best direction 
    for improvement.
    """)

    st.latex(r'''
    L = \sum_{i=1}^n l(y_i, \hat{y}_i) + \sum_{k=1}^K \Omega(f_k)
    ''')
    st.write("""
    Where:
    - L is the loss function
    - l is a differentiable convex loss function
    - Ω is a regularization term
    """)

    st.subheader("4. Feature Importance")
    st.write("""
    XGBoost provides a measure of feature importance based on how often a feature is used to split the data across all trees.
    """)

    # Create a simple diagram to illustrate XGBoost
    fig, ax = plt.subplots(figsize=(10, 6))
    ax.set_title("XGBoost: Sequential Tree Building")
    ax.set_xlabel("Features")
    ax.set_ylabel("Target")
    ax.scatter(np.random.rand(100), np.random.rand(100), alpha=0.5, label="Data points")
    
    for i in range(3):
        rect = plt.Rectangle((0.1 + i*0.3, 0.1), 0.2, 0.8, fill=False, label=f"Tree {i+1}")
        ax.add_patch(rect)
    
    ax.legend()
    st.pyplot(fig)

    st.header("Advantages of XGBoost")
    advantages = [
        "High performance and fast execution",
        "Handles missing values automatically",
        "Built-in regularization to prevent overfitting",
        "Supports parallel and distributed computing",
        "Flexibility (can solve regression, classification, and ranking problems)"
    ]
    for adv in advantages:
        st.write(f"- {adv}")

    st.header("Example: XGBoost in Action")
    st.code("""
    import xgboost as xgb
    from sklearn.datasets import make_regression
    from sklearn.model_selection import train_test_split

    # Create sample data
    X, y = make_regression(n_samples=1000, n_features=10, noise=0.1)
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2)

    # Create and train the model
    model = xgb.XGBRegressor(n_estimators=100, learning_rate=0.1)
    model.fit(X_train, y_train)

    # Make predictions
    predictions = model.predict(X_test)

    # Evaluate the model
    mse = mean_squared_error(y_test, predictions)
    print(f"Mean Squared Error: {mse}")

    # Feature importance
    importance = model.feature_importances_
    for i, imp in enumerate(importance):
        print(f"Feature {i} importance: {imp}")
    """, language="python")

    st.write("""
    This example demonstrates how to use XGBoost for a regression task, including model training, 
    prediction, evaluation, and examining feature importance.
    """)
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
            fig2 = create_advanced_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse)
            
            return fig, sept_achievement, lower_achievement, upper_achievement, rmse
        else:
            return None, None, None, None, None
    except Exception as e:
        st.error(f"Error in predict_and_visualize: {str(e)}")
        raise

import streamlit as st
import plotly.graph_objects as go
import plotly.subplots as sp
import pandas as pd
import numpy as np
from scipy import stats
def create_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse):
    fig = plt.figure(figsize=(20, 28))  # Increased height to accommodate new table
    gs = fig.add_gridspec(7, 2, height_ratios=[0.5, 0.5, 0.5, 3, 1, 2, 1])
    ax_region = fig.add_subplot(gs[0, :])
    ax_region.axis('off')
    ax_region.text(0.5, 0.5, f'{region}({brand})', fontsize=28, fontweight='bold', ha='center', va='center')
            
    # New table for current month sales data
    ax_current = fig.add_subplot(gs[1, :])
    ax_current.axis('off')
    current_data = [
                ['Total Sales\nTill Now','Trade %', 'Commitment\nfor Today', 'Asking\nfor Today', 'Yesterday\nSales', 'Yesterday\nCommitment'],
                [f"{region_data['Till Yesterday Total Sales'].iloc[-1]:.0f}",
                 f"{region_data['Till Yesterday Total sales in Trade'].iloc[-1]/region_data['Till Yesterday Total Sales'].iloc[-1]*100:.0f}",
                 f"{region_data['Commitment for Today'].iloc[-1]:.0f}",
                 f"{region_data['Asking for Today'].iloc[-1]:.0f}",
                 f"{region_data['Yesterday Sales'].iloc[-1]:.0f}",
                 f"{region_data['Yesterday Commitment'].iloc[-1]:.0f}"]
            ]
    current_table = ax_current.table(cellText=current_data[1:], colLabels=current_data[0], cellLoc='center', loc='center')
    current_table.auto_set_font_size(False)
    current_table.set_fontsize(10)
    current_table.scale(1, 1.7)
    for (row, col), cell in current_table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='black')
                    cell.set_facecolor('goldenrod')
                cell.set_edgecolor('brown')
            
            # Existing table (same as before)
    ax_table = fig.add_subplot(gs[2, :])
    ax_table.axis('off')
    table_data = [
                ['Month Target\n(Sep)', 'Monthly Achievement\n(Aug)', 'Predicted Achievement\n(Sept)(using XGBoost Algorithm)', 'CI', 'RMSE'],
                [f"{sept_target:.2f}", f"{region_data['Monthly Achievement(Aug)'].iloc[-1]:.2f}", 
                 f"{sept_achievement:.2f}", f"({lower_achievement:.2f}, {upper_achievement:.2f})", f"{rmse:.4f}"]
            ]
    table = ax_table.table(cellText=table_data[1:], colLabels=table_data[0], cellLoc='center', loc='center')
    table.auto_set_font_size(False)
    table.set_fontsize(10)
    table.scale(1, 1.7)
    for (row, col), cell in table.get_celld().items():
                if row == 0:
                    cell.set_text_props(fontweight='bold', color='black')
                    cell.set_facecolor('goldenrod')
                cell.set_edgecolor('brown')

    
    
    
    # Main bar chart (same as before)
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
        ax1.text(i, (max(target, achievement)+min(target,achievement))/2, f'{percentage:.1f}%', 
                 ha='center', va='bottom', fontsize=10, color=color, fontweight='bold')
    
    ax1.errorbar(x[-1] + width/2, sept_achievement, 
                 yerr=[[sept_achievement - lower_achievement], [upper_achievement - sept_achievement]],
                 fmt='o', color='darkred', capsize=5, capthick=2, elinewidth=2)
    
    # Percentage achievement line chart (same as before)
    ax2 = fig.add_subplot(gs[4, :])
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
    ax3 = fig.add_subplot(gs[5, :])
    ax3.axis('off')
    
    current_year = 2024  # Assuming the current year is 2024
    last_year = 2023

    channel_data = [
        ('Trade', region_data['Trade Aug'].iloc[-1], region_data['Trade Aug 2023'].iloc[-1]),
        ('Premium', region_data['Premium Aug'].iloc[-1], region_data['Premium Aug 2023'].iloc[-1]),
        ('Blended', region_data['Blended Aug'].iloc[-1], region_data['Blended Aug 2023'].iloc[-1])
    ]
    monthly_achievement_aug = region_data['Monthly Achievement(Aug)'].iloc[-1]
    total_aug_current = region_data['Monthly Achievement(Aug)'].iloc[-1]
    total_aug_last = region_data['Total Aug 2023'].iloc[-1]
    
    ax3.text(0.2, 1, f'\nAugust {current_year} Sales Breakdown:-', fontsize=16, fontweight='bold', ha='center', va='center')
    
    # Helper function to create arrow
    def get_arrow(value):
        return '↑' if value > 0 else '↓' if value < 0 else '→'

    # Helper function to get color
    def get_color(value):
        return 'green' if value > 0 else 'red' if value < 0 else 'black'

    # Display total sales
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

    

    
    # Updated: August Region Type Breakdown with values
    ax4 = fig.add_subplot(gs[5, 1])
    region_type_data = [
        region_data['Green Aug'].iloc[-1],
        region_data['Yellow Aug'].iloc[-1],
        region_data['Red Aug'].iloc[-1],
        region_data['Unidentified Aug'].iloc[-1]
    ]
    region_type_labels = ['Green', 'Yellow', 'Red', 'Unidentified']
    colors = ['green', 'yellow', 'red', 'gray']
    
    def make_autopct(values):
        def my_autopct(pct):
            total = sum(values)
            val = int(round(pct*total/100.0))
            return f'{pct:.1f}%\n({val:.0f})'
        return my_autopct
    
    ax4.pie(region_type_data, labels=region_type_labels, colors=colors,
            autopct=make_autopct(region_type_data), startangle=90)
    ax4.set_title('August 2024 Region Type Breakdown:-', fontsize=16, fontweight='bold')
    ax5 = fig.add_subplot(gs[6, :])
    ax5.axis('off')
    
    q3_table_data = [
        ['Overall Requirement', 'Requirement in\nTrade Channel', 'Requirement in\nBlednded Product Category', 'Requirement for\nPremium Product'],
        [f"{region_data['Q3 2023'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Trade'].iloc[-1]:.0f}", 
         f"{region_data['Q3 2023 Blended'].iloc[-1]:.0f}", f"{region_data['Q3 2023 Premium'].iloc[-1]:.0f}"]
    ]
    
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
    plt.tight_layout()
    return fig

def create_advanced_visualization(region_data, region, brand, months, sept_target, sept_achievement, lower_achievement, upper_achievement, rmse):
    st.title(f"{region} ({brand}) Performance")

    # Current Month Sales Data
    st.header("Current Month Sales Overview")
    col1, col2, col3 = st.columns(3)
    col1.metric("Total Sales Till Now", f"{region_data['Till Yesterday Total Sales'].iloc[-1]:,.0f}")
    col2.metric("Trade %", f"{region_data['Till Yesterday Total sales in Trade'].iloc[-1]/region_data['Till Yesterday Total Sales'].iloc[-1]*100:.1f}%")
    col3.metric("Commitment for Today", f"{region_data['Commitment for Today'].iloc[-1]:,.0f}")

    col4, col5, col6 = st.columns(3)
    col4.metric("Asking for Today", f"{region_data['Asking for Today'].iloc[-1]:,.0f}")
    col5.metric("Yesterday Sales", f"{region_data['Yesterday Sales'].iloc[-1]:,.0f}")
    col6.metric("Yesterday Commitment", f"{region_data['Yesterday Commitment'].iloc[-1]:,.0f}")

    # Monthly Performance Table
    st.header("Monthly Performance Summary")
    performance_df = pd.DataFrame({
        "Metric": ["Month Target (Sep)", "Monthly Achievement (Aug)", "Predicted Achievement (Sep)", "Confidence Interval", "RMSE"],
        "Value": [
            f"{sept_target:,.2f}",
            f"{region_data['Monthly Achievement(Aug)'].iloc[-1]:,.2f}",
            f"{sept_achievement:,.2f}",
            f"({lower_achievement:,.2f}, {upper_achievement:,.2f})",
            f"{rmse:.4f}"
        ]
    })
    st.table(performance_df)

    # Main Chart: Monthly Targets and Achievements
    st.header("Monthly Targets and Achievements for FY 2025")
    
    all_months = months + ['Sep']
    actual_achievements = [region_data[f'Monthly Achievement({month})'].iloc[-1] for month in months]
    actual_targets = [region_data[f'Month Tgt ({month})'].iloc[-1] for month in months]
    all_achievements = actual_achievements + [sept_achievement]
    all_targets = actual_targets + [sept_target]

    fig1 = go.Figure()
    fig1.add_trace(go.Bar(x=all_months, y=all_targets, name='Target', marker_color='rgba(255, 192, 203, 0.8)'))
    fig1.add_trace(go.Bar(x=all_months, y=all_achievements, name='Achievement', marker_color='rgba(255, 255, 0, 0.8)'))
    fig1.add_trace(go.Scatter(x=['Sep'], y=[sept_achievement], mode='markers', name='September Prediction',
                             marker=dict(color='red', size=10), error_y=dict(
                                 type='data',
                                 symmetric=False,
                                 array=[upper_achievement - sept_achievement],
                                 arrayminus=[sept_achievement - lower_achievement],
                                 visible=True
                             )))

    for i, (target, achievement) in enumerate(zip(all_targets, all_achievements)):
        percentage = (achievement / target) * 100
        color = 'green' if percentage >= 100 else 'red'
        fig1.add_annotation(x=all_months[i], y=max(target, achievement),
                           text=f'{percentage:.1f}%',
                           showarrow=False,
                           font=dict(size=10, color=color))

    fig1.update_layout(barmode='group', title='Monthly Targets and Achievements', 
                      xaxis_title='Month', yaxis_title='Value')
    st.plotly_chart(fig1, use_container_width=True)

    # Percentage Achievement Line Chart
    st.header("Percentage Achievement Trend")
    percent_achievements = [((ach / tgt) * 100) for ach, tgt in zip(all_achievements, all_targets)]
    
    fig_percent = go.Figure()
    fig_percent.add_trace(go.Scatter(x=all_months, y=percent_achievements, mode='lines+markers', name='% Achievement'))
    fig_percent.add_hline(y=100, line_dash="dash", line_color="red", annotation_text="100% Target")

    for i, pct in enumerate(percent_achievements):
        fig_percent.add_annotation(x=all_months[i], y=pct, text=f'{pct:.1f}%', showarrow=False, yshift=10)

    fig_percent.update_layout(title='Percentage Achievement by Month', xaxis_title='Month', yaxis_title='% Achievement')
    st.plotly_chart(fig_percent, use_container_width=True)

    # August Sales Breakdown
    st.header("August 2024 Sales Breakdown")
    col_breakdown1, col_breakdown2 = st.columns(2)

    with col_breakdown1:
        st.subheader("Channel Performance")
        channel_data = [
            ('Trade', region_data['Trade Aug'].iloc[-1], region_data['Trade Aug 2023'].iloc[-1]),
            ('Premium', region_data['Premium Aug'].iloc[-1], region_data['Premium Aug 2023'].iloc[-1]),
            ('Blended', region_data['Blended Aug'].iloc[-1], region_data['Blended Aug 2023'].iloc[-1])
        ]
        
        for channel, value_current, value_last in channel_data:
            change = ((value_current - value_last) / value_last) * 100
            st.metric(
                label=channel,
                value=f"{value_current:,.0f}",
                delta=f"{change:+.1f}% vs Last Year"
            )

    with col_breakdown2:
        st.subheader("Region Type Breakdown")
        region_type_data = [
            region_data['Green Aug'].iloc[-1],
            region_data['Yellow Aug'].iloc[-1],
            region_data['Red Aug'].iloc[-1],
            region_data['Unidentified Aug'].iloc[-1]
        ]
        region_type_labels = ['Green', 'Yellow', 'Red', 'Unidentified']
        colors = ['green', 'yellow', 'red', 'gray']

        fig_pie = go.Figure(data=[go.Pie(labels=region_type_labels, values=region_type_data, marker_colors=colors)])
        fig_pie.update_traces(textposition='inside', textinfo='percent+label')
        fig_pie.update_layout(title='August 2024 Region Type Distribution')
        st.plotly_chart(fig_pie, use_container_width=True)

    # Quarterly Requirements
    st.header("Quarterly Requirements for September 2024")
    q3_data = {
        'Category': ['Overall', 'Trade Channel', 'Blended Product', 'Premium Product'],
        'Requirement': [
            region_data['Q3 2023'].iloc[-1],
            region_data['Q3 2023 Trade'].iloc[-1],
            region_data['Q3 2023 Blended'].iloc[-1],
            region_data['Q3 2023 Premium'].iloc[-1]
        ]
    }
    q3_df = pd.DataFrame(q3_data)
    
    fig_q3 = go.Figure([go.Bar(x=q3_df['Category'], y=q3_df['Requirement'])])
    fig_q3.update_layout(title='Q3 Requirements by Category', xaxis_title='Category', yaxis_title='Requirement')
    st.plotly_chart(fig_q3, use_container_width=True)

    # Additional Insights
    st.header("Additional Insights")
    total_aug_current = region_data['Monthly Achievement(Aug)'].mean()  # or use .sum() as needed
    total_aug_last = region_data['Total Aug 2023'].mean()  # or use .sum() as needed
    yoy_growth = ((total_aug_current - total_aug_last) / total_aug_last) * 100

    # Now this should work without raising an error
    st.metric("Year-over-Year Growth", f"{yoy_growth:.1f}%", delta=f"{yoy_growth:.1f}%")



    # Trend Analysis
    trend_data = region_data[[f'Monthly Achievement({month})' for month in months]].iloc[-1].values
    trend, _ = np.polyfit(range(len(trend_data)), trend_data, 1)
    trend_direction = "Upward" if trend > 0 else "Downward" if trend < 0 else "Stable"
    st.write(f"Sales Trend: {trend_direction}")

    # Statistical Significance
    _, p_value = stats.ttest_ind(actual_achievements, actual_targets)
    st.write(f"Statistical Significance (p-value): {p_value:.4f}")
    if p_value < 0.05:
        st.write("There is a statistically significant difference between achievements and targets.")
    else:
        st.write("There is no statistically significant difference between achievements and targets.")

    # Recommendations
    st.subheader("Recommendations")
    if sept_achievement < sept_target:
        st.write("- Focus on increasing sales in underperforming channels")
        st.write("- Implement targeted marketing strategies for regions with lower performance")
    else:
        st.write("- Maintain current strategies and look for opportunities to exceed targets")
        st.write("- Share best practices from high-performing regions with others")

    st.write("- Continue monitoring market trends and adjust forecasts accordingly")
    st.write("- Regularly review and optimize the sales pipeline")

    return None  # No need to return a figure, as we're using Streamlit components directly

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
                        
                        main_table_data.append([
                            region, brand, f"{sept_target:.0f}", f"{aug_achievement:.0f}",
                            f"{sept_achievement:.0f}", f"({lower_achievement:.2f},\n{upper_achievement:.2f})", f"{rmse:.4f}"
                        ])
                        
                        additional_table_data.append([
                            region, brand, 
                            f"{region_data['Till Yesterday Total Sales'].iloc[-1]:.0f}",
                            f"{region_data['Commitment for Today'].iloc[-1]:.0f}",
                            f"{region_data['Asking for Today'].iloc[-1]:.0f}",
                            f"{region_data['Yesterday Sales'].iloc[-1]:.0f}",
                            f"{region_data['Yesterday Commitment'].iloc[-1]:.0f}"
                        ])
                        
                        valid_data = True
                    else:
                        st.warning(f"No data available for {region} and {brand}")
            except Exception as e:
                st.warning(f"Error processing {region} and {brand}: {str(e)}")
    
    if valid_data:
        num_rows = len(main_table_data) + len(additional_table_data)
        fig_height = max(12, 2 + 0.5 * num_rows)  # Increased minimum height
        
        fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, fig_height), gridspec_kw={'height_ratios': [1, 1.5]})
        fig.suptitle("", fontsize=16, fontweight='bold', y=0.98)
        
        # Function to create styled table
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
        
        # Create additional table
        create_styled_table(ax1, additional_table_data, "Current Month Sales Data")
        
        # Create main table
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
def sales_prediction_app():
    st.title("📊 Sales Prediction App")
    
    # Load Lottie animation
    lottie_url = "https://assets5.lottiefiles.com/packages/lf20_V9t630.json"
    lottie_json = load_lottie_url(lottie_url)
    
    # Initialize session state variables if they don't exist
    if 'df' not in st.session_state:
        st.session_state['df'] = None
    if 'regions' not in st.session_state:
        st.session_state['regions'] = []
    if 'brands' not in st.session_state:
        st.session_state['brands'] = []
    
    # Sidebar
    with st.sidebar:
        st_lottie(lottie_json, height=200)
        st.title("Navigation")
        page = st.radio("Go to", ["Home", "Predictions", "XGBoost Explained", "About"])
    
    if page == "Home":
        st.write("This app helps you predict and visualize sales achievements for different regions and brands.")
        st.write("Use the sidebar to navigate between pages and upload your data to get started!")
        
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key="Sales_Prediction_uploader")
        if uploaded_file is not None:
            with st.spinner("Loading data..."):
                df, regions, brands = load_data(uploaded_file)
            st.session_state['df'] = df
            st.session_state['regions'] = regions
            st.session_state['brands'] = brands
            st.success("File uploaded and processed successfully!")
    
    elif page == "Predictions":
        st.subheader("🔮 Sales Predictions")
        if st.session_state['df'] is None:
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
                    st.error("Unable to generate combined report. Please check the warnings above for more details.")

    elif page == "XGBoost Explained":
        xgboost_explanation()
    
    elif page == "About":
        st.subheader("ℹ️ About the Sales Prediction App")
        st.write("""
        This app is designed to help sales teams predict and visualize their performance across different regions and brands.
        
        Key features:
        - Data upload and processing
        - Individual predictions for each region and brand
        - Combined report generation
        - Interactive visualizations
        
        For any questions or support, please contact our team at support@salespredictionapp.com
        """)
def load_lottie_url(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def generate_shareable_link(file_path):
    file_name = os.path.basename(file_path)
    encoded_file_name = quote(file_name)
    return f"https://your-file-sharing-service.com/files/{encoded_file_name}"

def get_online_editor_url(file_extension):
    extension_mapping = {
        '.xlsx': 'https://www.office.com/launch/excel?auth=2',
        '.xls': 'https://www.office.com/launch/excel?auth=2',
        '.doc': 'https://www.office.com/launch/word?auth=2',
        '.docx': 'https://www.office.com/launch/word?auth=2',
        '.ppt': 'https://www.office.com/launch/powerpoint?auth=2',
        '.pptx': 'https://www.office.com/launch/powerpoint?auth=2',
        '.pdf': 'https://documentcloud.adobe.com/link/home/'
    }
    return extension_mapping.get(file_extension.lower(), 'https://www.google.com/drive/')

def folder_menu():
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
    .file-box {
        border: 1px solid #ddd;
        padding: 15px;
        margin: 15px 0;
        border-radius: 10px;
        background-color: #f9f9f9;
        transition: all 0.3s ease;
    }
    .file-box:hover {
        box-shadow: 0 5px 15px rgba(0,0,0,0.1);
        transform: translateY(-5px);
    }
    .stButton>button {
        border-radius: 20px;
        padding: 10px 20px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        transform: scale(1.05);
    }
    .upload-section {
        background-color: #e6f3ff;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .todo-section {
        background-color: #f0f8ff;
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
    }
    .todo-item {
        display: flex;
        align-items: center;
        margin-bottom: 10px;
    }
    .todo-text {
        margin-left: 10px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>Advanced File Manager</span></div>', unsafe_allow_html=True)

    # Load Lottie animation
    lottie_urls = [
        "https://assets9.lottiefiles.com/packages/lf20_3vbOcw.json",  # File manager animation
        "https://assets9.lottiefiles.com/packages/lf20_5lAtR7.json",  # Folder animation
        "https://assets1.lottiefiles.com/packages/lf20_4djadwfo.json",  # Document management
        "https://assets6.lottiefiles.com/packages/lf20_2a5yxpci.json"   # File transfer
    ]

    # Try loading Lottie animations
    lottie_json = None
    for url in lottie_urls:
        lottie_json = load_lottie_url(url)
        if lottie_json:
            break
    
    col1, col2 = st.columns([1, 2])
    with col1:
        if lottie_json:
           st_lottie(lottie_json, height=200, key="file_animation")
        else:
           st.image("https://via.placeholder.com/200x200.png?text=File+Manager", use_column_width=True)
    with col2:
        st.markdown("""
        Welcome to the Advanced File Manager! 
        Here you can upload, download, and manage your files with ease. 
        Enjoy the smooth animations, user-friendly interface, and new features like file search and sorting.
        """)

    # Create a folder to store uploaded files if it doesn't exist
    if not os.path.exists("uploaded_files"):
        os.makedirs("uploaded_files")

    # File uploader
    st.markdown('<div class="upload-section">', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Upload a file", type=["xlsx", "xls", "doc", "docx", "pdf", "ppt", "pptx", "txt", "csv"])
    if uploaded_file is not None:
        file_details = {"FileName": uploaded_file.name, "FileType": uploaded_file.type, "FileSize": uploaded_file.size}
        
        # Save the uploaded file
        with open(os.path.join("uploaded_files", uploaded_file.name), "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"File {uploaded_file.name} saved successfully!")
    st.markdown('</div>', unsafe_allow_html=True)

    # Display uploaded files
    st.subheader("Your Files")
    
    # File search and sorting
    search_query = st.text_input("Search files", "")
    sort_option = st.selectbox("Sort by", ["Name", "Size", "Date Modified"])

    # Use session state to track file deletion
    if 'files_to_delete' not in st.session_state:
        st.session_state.files_to_delete = set()

    files = os.listdir("uploaded_files")
    
    # Apply search filter
    if search_query:
        files = [f for f in files if search_query.lower() in f.lower()]
    
    # Apply sorting
    if sort_option == "Name":
        files.sort()
    elif sort_option == "Size":
        files.sort(key=lambda x: os.path.getsize(os.path.join("uploaded_files", x)), reverse=True)
    elif sort_option == "Date Modified":
        files.sort(key=lambda x: os.path.getmtime(os.path.join("uploaded_files", x)), reverse=True)

    for filename in files:
        file_path = os.path.join("uploaded_files", filename)
        file_stats = os.stat(file_path)
        
        st.markdown(f'<div class="file-box">', unsafe_allow_html=True)
        col1, col2, col3, col4, col5 = st.columns([3, 1, 1, 1, 1])
        with col1:
            st.markdown(f"<h3>{filename}</h3>", unsafe_allow_html=True)
            st.text(f"Size: {file_stats.st_size / 1024:.2f} KB")
            st.text(f"Modified: {datetime.fromtimestamp(file_stats.st_mtime).strftime('%Y-%m-%d %H:%M:%S')}")
        with col2:
            if st.button(f"📥 Download", key=f"download_{filename}"):
                with open(file_path, "rb") as file:
                    file_content = file.read()
                    b64 = base64.b64encode(file_content).decode()
                    href = f'<a href="data:application/octet-stream;base64,{b64}" download="{filename}">Click to download</a>'
                    st.markdown(href, unsafe_allow_html=True)
        with col3:
            if st.button(f"🗑️ Delete", key=f"delete_{filename}"):
                st.session_state.files_to_delete.add(filename)
        with col4:
            file_extension = os.path.splitext(filename)[1]
            editor_url = get_online_editor_url(file_extension)
            st.markdown(f"[🌐 Open Online]({editor_url})")
        with col5:
            shareable_link = generate_shareable_link(file_path)
            st.markdown(f"[🔗 Share]({shareable_link})")
        st.markdown('</div>', unsafe_allow_html=True)

    # Process file deletion
    files_deleted = False
    for filename in st.session_state.files_to_delete:
        file_path = os.path.join("uploaded_files", filename)
        if os.path.exists(file_path):
            os.remove(file_path)
            st.warning(f"{filename} has been deleted.")
            files_deleted = True
    
    # Clear the set of files to delete
    st.session_state.files_to_delete.clear()

    # Rerun the app if any files were deleted
    if files_deleted:
        st.rerun()

    st.info("Note: The 'Open Online' links will redirect you to the appropriate online editor. You may need to manually open your file once there.")

    # To-Do List / Diary Section
    st.markdown('<div class="todo-section">', unsafe_allow_html=True)
    st.subheader("📝 To-Do List / Diary")

    # Load existing to-do items
    if 'todo_items' not in st.session_state:
        st.session_state.todo_items = []

    # Add new to-do item
    new_item = st.text_input("Add a new to-do item or diary entry")
    if st.button("Add"):
        if new_item:
            st.session_state.todo_items.append({"text": new_item, "done": False, "date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")})
            st.success("Item added successfully!")

    # Display and manage to-do items
    for idx, item in enumerate(st.session_state.todo_items):
        col1, col2, col3 = st.columns([0.1, 3, 1])
        with col1:
            done = st.checkbox("", item["done"], key=f"todo_{idx}")
            if done != item["done"]:
                st.session_state.todo_items[idx]["done"] = done
        with col2:
            st.markdown(f"<div class='todo-text'>{'<s>' if item['done'] else ''}{item['text']}{'</s>' if item['done'] else ''}</div>", unsafe_allow_html=True)
        with col3:
            st.text(item["date"])
        if st.button("Delete", key=f"delete_todo_{idx}"):
            st.session_state.todo_items.pop(idx)
            st.rerun()

    st.markdown('</div>', unsafe_allow_html=True)

    # Add a fun fact section
    st.markdown("---")
    st.subheader("📚 Fun File Fact")
    fun_facts = [
        "The first computer virus was created in 1983 and was called the Elk Cloner.",
        "The most common file extension in the world is .dll (Dynamic Link Library).",
        "The largest file size theoretically possible in Windows is 16 exabytes minus 1 KB.",
        "The PDF file format was invented by Adobe in 1993.",
        "The first widely-used image format on the web was GIF, created in 1987."
    ]
    st.markdown(f"*{fun_facts[int(os.urandom(1)[0]) % len(fun_facts)]}*")
def load_lottieurl(url: str):
    r = requests.get(url)
    if r.status_code != 200:
        return None
    return r.json()

def sales_dashboard():
    
    st.title("Sales Dashboard")

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
    .section-box {
        background-color: #f9f9f9;
        border-radius: 10px;
        padding: 20px;
        margin-bottom: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        transition: all 0.3s ease;
    }
    .section-box:hover {
        transform: translateY(-5px);
        box-shadow: 0 6px 8px rgba(0, 0, 0, 0.15);
    }
    .upload-section {
        background-color: #e6f3ff;
        padding: 20px;
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .stDataFrame {
        font-family: 'Arial', sans-serif;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="title"><span>Welcome to SalesDashboard</span></div>', unsafe_allow_html=True)

    # Load Lottie animation
    lottie_url = "https://assets2.lottiefiles.com/packages/lf20_V9t630.json"  # New interesting animation
    lottie_json = load_lottie_url(lottie_url)
    
    col1, col2 = st.columns([1, 2])
    with col1:
        st_lottie(lottie_json, height=200, key="home_animation")
    with col2:
        st.markdown("""
        Welcome to our interactive Sales Analysis Dashboard! 
        This powerful tool helps you analyze Sales data for JKLC and UCWL across different regions, districts and channels.
        Let's get started with your data analysis journey!
        """)

    # File upload
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key="Sales_Dashboard_uploader")
    st.markdown('<div class="section-box">', unsafe_allow_html=True)
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        df = process_dataframe(df)

        # Region selection
        regions = df['Region'].unique()
        selected_regions = st.multiselect('Select Regions', regions)

        # District selection
        districts = df[df['Region'].isin(selected_regions)]['Dist Name'].unique()
        selected_districts = st.multiselect('Select Districts', districts)

        # Channel selection
        channels = ['Overall', 'Trade', 'Non-Trade']
        selected_channels = st.multiselect('Select Channels', channels, default=channels)

        # Checkbox for whole region totals
        show_whole_region = st.checkbox('Show whole region totals')

        if st.button('Generate Report'):
            display_data(df, selected_regions, selected_districts, selected_channels, show_whole_region)

def process_dataframe(df):
    
    
    column_mapping = {
    pd.to_datetime('2024-09-23 00:00:00'): '23-Sep',
    pd.to_datetime('2024-08-23 00:00:00'): '23-Aug',
    pd.to_datetime('2024-07-23 00:00:00'): '23-Jul',
    pd.to_datetime('2024-06-23 00:00:00'): '23-Jun',
    pd.to_datetime('2024-05-23 00:00:00'): '23-May',
    pd.to_datetime('2024-04-23 00:00:00'): '23-Apr',
    pd.to_datetime('2024-08-24 00:00:00'): '24-Aug',
    pd.to_datetime('2024-07-24 00:00:00'): '24-Jul',
    pd.to_datetime('2024-06-24 00:00:00'): '24-Jun',
    pd.to_datetime('2024-05-24 00:00:00'): '24-May',
    pd.to_datetime('2024-04-24 00:00:00'): '24-Apr'
}

    df = df.rename(columns=column_mapping)
    df['FY 2024 till Aug'] = df['24-Apr'] + df['24-May'] + df['24-Jun'] + df['24-Jul'] + df['24-Aug']
    df['FY 2023 till Aug'] = df['23-Apr'] + df['23-May'] + df['23-Jun'] + df['23-Jul'] + df['23-Aug']
    df['Quarterly Requirement'] = df['23-Jul'] + df['23-Aug'] + df['23-Sep'] - df['24-Jul'] - df['24-Aug']
    df['Growth/Degrowth(MTD)'] = (df['24-Aug'] - df['23-Aug']) / df['23-Aug'] * 100
    df['Growth/Degrowth(YTD)'] = (df['FY 2024 till Aug'] - df['FY 2023 till Aug']) / df['FY 2023 till Aug'] * 100
    df['Q3 2023'] = df['23-Jul'] + df['23-Aug'] + df['23-Sep']
    df['Q3 2024 till August'] = df['24-Jul'] + df['24-Aug']

    # Non-Trade calculations
    for month in ['Sep', 'Aug', 'Jul', 'Jun', 'May', 'Apr']:
        df[f'23-{month} Non-Trade'] = df[f'23-{month}'] - df[f'23-{month} Trade']
        if month != 'Sep':
            df[f'24-{month} Non-Trade'] = df[f'24-{month}'] - df[f'24-{month} Trade']

    # Trade calculations
    df['FY 2024 till Aug Trade'] = df['24-Apr Trade'] + df['24-May Trade'] + df['24-Jun Trade'] + df['24-Jul Trade'] + df['24-Aug Trade']
    df['FY 2023 till Aug Trade'] = df['23-Apr Trade'] + df['23-May Trade'] + df['23-Jun Trade'] + df['23-Jul Trade'] + df['23-Aug Trade']
    df['Quarterly Requirement Trade'] = df['23-Jul Trade'] + df['23-Aug Trade'] + df['23-Sep Trade'] - df['24-Jul Trade'] - df['24-Aug Trade']
    df['Growth/Degrowth(MTD) Trade'] = (df['24-Aug Trade'] - df['23-Aug Trade']) / df['23-Aug Trade'] * 100
    df['Growth/Degrowth(YTD) Trade'] = (df['FY 2024 till Aug Trade'] - df['FY 2023 till Aug Trade']) / df['FY 2023 till Aug Trade'] * 100
    df['Q3 2023 Trade'] = df['23-Jul Trade'] + df['23-Aug Trade'] + df['23-Sep Trade']
    df['Q3 2024 till August Trade'] = df['24-Jul Trade'] + df['24-Aug Trade']

    # Non-Trade calculations
    df['FY 2024 till Aug Non-Trade'] = df['24-Apr Non-Trade'] + df['24-May Non-Trade'] + df['24-Jun Non-Trade'] + df['24-Jul Non-Trade'] + df['24-Aug Non-Trade']
    df['FY 2023 till Aug Non-Trade'] = df['23-Apr Non-Trade'] + df['23-May Non-Trade'] + df['23-Jun Non-Trade'] + df['23-Jul Non-Trade'] + df['23-Aug Non-Trade']
    df['Quarterly Requirement Non-Trade'] = df['23-Jul Non-Trade'] + df['23-Aug Non-Trade'] + df['23-Sep Non-Trade'] - df['24-Jul Non-Trade'] - df['24-Aug Non-Trade']
    df['Growth/Degrowth(MTD) Non-Trade'] = (df['24-Aug Non-Trade'] - df['23-Aug Non-Trade']) / df['23-Aug Non-Trade'] * 100
    df['Growth/Degrowth(YTD) Non-Trade'] = (df['FY 2024 till Aug Non-Trade'] - df['FY 2023 till Aug Non-Trade']) / df['FY 2023 till Aug Non-Trade'] * 100
    df['Q3 2023 Non-Trade'] = df['23-Jul Non-Trade'] + df['23-Aug Non-Trade'] + df['23-Sep Non-Trade']
    df['Q3 2024 till August Non-Trade'] = df['24-Jul Non-Trade'] + df['24-Aug Non-Trade']

    # Handle division by zero

    return df

    pass
def display_data(df, selected_regions, selected_districts, selected_channels, show_whole_region):
    def color_growth(val):
        try:
            value = float(val.strip('%'))
            color = 'green' if value > 0 else 'red' if value < 0 else 'black'
            return f'color: {color}'
        except:
            return 'color: black'

    if show_whole_region:
        filtered_data = df[df['Region'].isin(selected_regions)].copy()
        
        # Calculate sums for relevant columns first
        sum_columns = ['24-Apr','24-May','24-Jun','24-Jul','24-Aug','23-Apr','23-May','23-Jun','23-Jul', '23-Aug', 'FY 2024 till Aug', 'FY 2023 till Aug', 'Q3 2023', 'Q3 2024 till August','24-Apr Trade','24-May Trade','24-Jun Trade','24-Jul Trade', 
                        '24-Aug Trade','23-Apr Trade','23-May Trade','23-Jun Trade','23-Jul Trade', '23-Aug Trade', 'FY 2024 till Aug Trade', 'FY 2023 till Aug Trade', 
                        'Q3 2023 Trade', 'Q3 2024 till August Trade','24-Apr Non-Trade','24-May Non-Trade','24-Jun Non-Trade','24-Jul Non-Trade',
                        '24-Aug Non-Trade','23-Apr Non-Trade','23-May Non-Trade','23-Jun Non-Trade','23-Jul Non-Trade', '23-Aug Non-Trade', 'FY 2024 till Aug Non-Trade', 'FY 2023 till Aug Non-Trade', 
                        'Q3 2023 Non-Trade', 'Q3 2024 till August Non-Trade']
        grouped_data = filtered_data.groupby('Region')[sum_columns].sum().reset_index()

        # Then calculate Growth/Degrowth based on the summed values
        grouped_data['Growth/Degrowth(MTD)'] = (grouped_data['24-Aug'] - grouped_data['23-Aug']) / grouped_data['23-Aug'] * 100
        grouped_data['Growth/Degrowth(YTD)'] = (grouped_data['FY 2024 till Aug'] - grouped_data['FY 2023 till Aug']) / grouped_data['FY 2023 till Aug'] * 100
        grouped_data['Quarterly Requirement'] = grouped_data['Q3 2023'] - grouped_data['Q3 2024 till August']

        grouped_data['Growth/Degrowth(MTD) Trade'] = (grouped_data['24-Aug Trade'] - grouped_data['23-Aug Trade']) / grouped_data['23-Aug Trade'] * 100
        grouped_data['Growth/Degrowth(YTD) Trade'] = (grouped_data['FY 2024 till Aug Trade'] - grouped_data['FY 2023 till Aug Trade']) / grouped_data['FY 2023 till Aug Trade'] * 100
        grouped_data['Quarterly Requirement Trade'] = grouped_data['Q3 2023 Trade'] - grouped_data['Q3 2024 till August Trade']

        grouped_data['Growth/Degrowth(MTD) Non-Trade'] = (grouped_data['24-Aug Non-Trade'] - grouped_data['23-Aug Non-Trade']) / grouped_data['23-Aug Non-Trade'] * 100
        grouped_data['Growth/Degrowth(YTD) Non-Trade'] = (grouped_data['FY 2024 till Aug Non-Trade'] - grouped_data['FY 2023 till Aug Non-Trade']) / grouped_data['FY 2023 till Aug Non-Trade'] * 100
        grouped_data['Quarterly Requirement Non-Trade'] = grouped_data['Q3 2023 Non-Trade'] - grouped_data['Q3 2024 till August Non-Trade']
    else:
        if selected_districts:
            filtered_data = df[df['Dist Name'].isin(selected_districts)].copy()
        else:
            filtered_data = df[df['Region'].isin(selected_regions)].copy()
        grouped_data = filtered_data

    for selected_channel in selected_channels:
        if selected_channel == 'Trade':
            columns_to_display = ['24-Aug Trade','23-Aug Trade','Growth/Degrowth(MTD) Trade','FY 2024 till Aug Trade', 'FY 2023 till Aug Trade','Growth/Degrowth(YTD) Trade','Q3 2023 Trade','Q3 2024 till August Trade', 'Quarterly Requirement Trade']
            suffix = ' Trade'
        elif selected_channel == 'Non-Trade':
            columns_to_display = ['24-Aug Non-Trade','23-Aug Non-Trade','Growth/Degrowth(MTD) Non-Trade','FY 2024 till Aug Non-Trade', 'FY 2023 till Aug Non-Trade','Growth/Degrowth(YTD) Non-Trade','Q3 2023 Non-Trade','Q3 2024 till August Non-Trade', 'Quarterly Requirement Non-Trade']
            suffix = ' Non-Trade'
        else:  # Overall
            columns_to_display = ['24-Aug','23-Aug','Growth/Degrowth(MTD)','FY 2024 till Aug', 'FY 2023 till Aug','Growth/Degrowth(YTD)','Q3 2023','Q3 2024 till August', 'Quarterly Requirement']
            suffix = ''
        
        display_columns = ['Region' if show_whole_region else 'Dist Name'] + columns_to_display
        
        st.subheader(f"{selected_channel} Sales Data")
        
        # Create a copy of the dataframe with only the columns we want to display
        display_df = grouped_data[display_columns].copy()
        
        # Set the 'Region' or 'Dist Name' column as the index
        display_df.set_index('Region' if show_whole_region else 'Dist Name', inplace=True)
        
        # Style the dataframe
        styled_df = display_df.style.format({
            col: '{:,.0f}' if 'Growth' not in col else '{:.2f}%' for col in columns_to_display
        }).applymap(color_growth, subset=[col for col in columns_to_display if 'Growth' in col])
        
        st.dataframe(styled_df)

        # Add a bar chart for YTD comparison
        fig = go.Figure(data=[
            go.Bar(name='FY 2023', x=grouped_data['Region' if show_whole_region else 'Dist Name'], y=grouped_data[f'FY 2023 till Aug{suffix}']),
            go.Bar(name='FY 2024', x=grouped_data['Region' if show_whole_region else 'Dist Name'], y=grouped_data[f'FY 2024 till Aug{suffix}']),
        ])
        fig.update_layout(barmode='group', title=f'{selected_channel} YTD Comparison')
        st.plotly_chart(fig)

        # Add a line chart for monthly trends including September 2024
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug']
        fig_trend = go.Figure()
        for year in ['23', '24']:
            y_values = []
            for month in months:
                column_name = f'{year}-{month}{suffix}'
                if column_name in grouped_data.columns:
                    y_values.append(grouped_data[column_name].sum())
                else:
                    y_values.append(None)
            
            fig_trend.add_trace(go.Scatter(
                x=months, 
                y=y_values, 
                mode='lines+markers+text',
                name=f'FY 20{year}',
                text=[f'{y:,.0f}' if y is not None else '' for y in y_values],
                textposition='top center'
            ))
        
        fig_trend.update_layout(
            title=f'{selected_channel} Monthly Trends', 
            xaxis_title='Month', 
            yaxis_title='Sales',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        st.plotly_chart(fig_trend)



def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def normal():
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"Product Mix Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "Product Mix Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the Product Mix Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Normal
            [random.randint(1500, 2500) for _ in range(12)],  # Premium
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.blue
        lc.lines[2].strokeColor = colors.pink
        lc.lines[3].strokeColor = colors.brown

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Normal EBITDA'),
            (colors.blue, 'Premium EBITDA'),
            (colors.crimson, 'Overall EBITDA'),
            (colors.brown, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Normal and Premium EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If both (Trade,Non-Trade) are present: Premium +5%, Normal -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Normal and Premium Product for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Premium Product , which typically have higher EBITDA.",
            "Analyze factors contributing to higher EBITDA in Premium Channel,and apply insights to Normal.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all channels.",
            "Invest in product innovation to expand Premium Product offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for Normal and Premium Product ceteris paribus.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across products,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Normal EBITDA', 'Premium EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Normal'] + filtered_df['Premium']
                    filtered_df[overall_col] = (
                        (filtered_df['Normal'] * filtered_df['Normal EBITDA'] +
                         filtered_df['Premium'] * filtered_df['Premium EBITDA'])/ total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Normal Share'] = filtered_df['Normal'] / total_quantity
                    filtered_df['Average Premium Share'] = filtered_df['Premium'] / total_quantity
                    
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        normal = row['Average Normal Share']
                        premium = row['Average Premium Share']
                        
                        if normal == 1 or premium == 1 :
                            # If any share is 100%, don't change
                            return normal,premium
                        else:
                            premium = min(premium + 0.05, 1)
                            normal = max(normal - 0.05, 1 - premium)
                        
                        return normal,premium
                    filtered_df['Adjusted Normal Share'], filtered_df['Adjusted Premium Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Normal Share'] * filtered_df['Normal EBITDA'] +
                        filtered_df['Adjusted Premium Share'] * filtered_df['Premium EBITDA']
                    )

                    # Calculate differences
                    filtered_df['P-N Difference'] = filtered_df['Premium EBITDA'] - filtered_df['Normal EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Normal EBITDA'],
                                             mode='lines+markers', name='Normal EBITDA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Premium EBITDA'],
                                             mode='lines+markers', name='Premium EBITDA', line=dict(color='blue')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='crimson', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBITDA',
                                             line=dict(color='brown', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(P-N: {g_r:.0f})<br>(I-O: {g_y:.0f}))" 
                                for month, g_r, g_y in 
                                zip(filtered_df['Month'], 
                                    filtered_df['P-N Difference'],  
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}(Type:-{product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Normal','Premium']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [45,45,45] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                    
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Normal Share', 'Average Premium Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Normal Share': 'green', 'Average Premium Share': 'blue'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Normal', 'Premium']]
                    for _, row in filtered_df[['Month', 'Normal', 'Premium','Average Normal Share', 'Average Premium Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Normal']:.0f} ({row['Average Normal Share']:.2%})",
                            f"{row['Premium']:.0f} ({row['Average Premium Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


 if selected == "Home":
    st.title("📊 Advanced Product Mix Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key='NormalvsPremiumuploader')
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
    st.title("📈 Product Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                subset_df = df[(df['Region'] == region) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Product_Mix_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Product_Mix_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type
        premium_share = st.sidebar.slider("Adjust Premium Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
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
            filtered_df[overall_col] = (filtered_df['Normal'] * filtered_df[cols[0]] +
                                        filtered_df['Premium'] * filtered_df[cols[1]]) / (
                                            filtered_df['Normal'] + filtered_df['Premium'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - premium_share/100) * filtered_df[cols[0]] +
                                          (premium_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            
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
            total_quantity = filtered_df['Normal'] + filtered_df['Premium']
            normal_share = (filtered_df['Normal'] / total_quantity * 100).round(2)
            premium_share = (filtered_df['Premium'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Premium Share (%)': premium_share,
                'Normal Share (%)': normal_share
            })
                  
            fig_pie = px.pie(share_df, values=[normal_share.mean(), premium_share.mean()], 
                                     names=['Normal', 'Premium'], title='Average Share Distribution',color=["N","P"],color_discrete_map={"N":"green","P":"blue"},hole=0.5)
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
    """)
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def trade():
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"Segment Mix Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "Segment Mix Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the Segment Mix Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Trade
            [random.randint(1500, 2500) for _ in range(12)],  # Non-Trade
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.blue
        lc.lines[2].strokeColor = colors.pink
        lc.lines[3].strokeColor = colors.brown

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Trade EBITDA'),
            (colors.blue, 'Non-Trade EBITDA'),
            (colors.crimson, 'Overall EBITDA'),
            (colors.brown, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Trade and Non-Trade EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If both (Trade,Non-Trade) are present: Trade +5%, Non-Trade -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Trade and Non-Trade Channel for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Trade Channel specifically for PPC, which typically have higher EBIDTA.",
            "Analyze factors contributing to higher EBIDTA in Trade Channel,and apply insights to Non-Trade.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all channels.",
            "Invest in product innovation to expand Trade Channel offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for Trade and Non-Trade channel ceteris paribus.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across GYR,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
                    filtered_df[overall_col] = (
                        (filtered_df['Trade'] * filtered_df['Trade EBITDA'] +
                         filtered_df['Non-Trade'] * filtered_df['Non-Trade EBITDA'])/ total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Trade Share'] = filtered_df['Trade'] / total_quantity
                    filtered_df['Average Non-Trade Share'] = filtered_df['Non-Trade'] / total_quantity
                    
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        trade = row['Average Trade Share']
                        nontrade = row['Average Non-Trade Share']
                        
                        if trade == 1 or nontrade == 1 :
                            # If any share is 100%, don't change
                            return trade,nontrade
                        else:
                            trade = min(trade + 0.05, 1)
                            nontrade = min(nontrade - 0.05, 1 - trade)
                        
                        return trade,nontrade
                    filtered_df['Adjusted Trade Share'], filtered_df['Adjusted Non-Trade Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Trade Share'] * filtered_df['Trade EBITDA'] +
                        filtered_df['Adjusted Non-Trade Share'] * filtered_df['Non-Trade EBITDA']
                    )

                    # Calculate differences
                    filtered_df['T-NT Difference'] = filtered_df['Trade EBITDA'] - filtered_df['Non-Trade EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Trade EBITDA'],
                                             mode='lines+markers', name='Trade EBIDTA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Non-Trade EBITDA'],
                                             mode='lines+markers', name='Non-Trade EBIDTA', line=dict(color='blue')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='crimson', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBIDTA',
                                             line=dict(color='brown', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(T-NT: {g_r:.0f})<br>(I-O: {g_y:.0f}))" 
                                for month, g_r, g_y in 
                                zip(filtered_df['Month'], 
                                    filtered_df['T-NT Difference'],  
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}(Type:-{product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Trade','Non-Trade']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [45,45,45] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                    
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Trade Share', 'Average Non-Trade Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Trade Share': 'green', 'Average Non-Trade Share': 'blue'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Trade', 'Non-Trade']]
                    for _, row in filtered_df[['Month', 'Trade', 'Non-Trade','Average Trade Share', 'Average Non-Trade Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Trade']:.0f} ({row['Average Trade Share']:.2%})",
                            f"{row['Non-Trade']:.0f} ({row['Average Non-Trade Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


 if selected == "Home":
    st.title("📊 Advanced Segment Mix Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key='TradevsNontradeuploader')
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
    st.title("📈 Segment Mix Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                subset_df = df[(df['Region'] == region) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="Segment_Mix_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset) & (df['Type'] != 'PPC Premium')]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type
        trade_share = st.sidebar.slider("Adjust Trade Share (%)", 0, 100, 50)

        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Trade NSR', 'Non-Trade NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Trade Contribution', 'Non-Trade Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Trade EBITDA', 'Non-Trade EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Trade'] * filtered_df[cols[0]] +
                                        filtered_df['Non-Trade'] * filtered_df[cols[1]]) / (
                                            filtered_df['Trade'] + filtered_df['Non-Trade'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - trade_share/100) * filtered_df[cols[1]] +
                                          (trade_share/100) * filtered_df[cols[0]])
            
            # Calculate difference between Premium and Normal
            filtered_df['Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            
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
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({trade_share}% Trade)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(T-NT: {diff:.2f})<br>(I-O: {i_diff:.2f})" for month, diff, i_diff in 
                        zip(filtered_df['Month'], filtered_df['Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (T-NT: Trade - Non-Trade, I-O: Imaginary - Overall)',
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
            st.subheader("Share of Trade and Non-Trade Channel")
            total_quantity = filtered_df['Trade'] + filtered_df['Non-Trade']
            trade_share = (filtered_df['Trade'] / total_quantity * 100).round(2)
            nontrade_share = (filtered_df['Non-Trade'] / total_quantity * 100).round(2)
            
            share_df = pd.DataFrame({
                'Month': filtered_df['Month'],
                'Trade Share (%)': trade_share,
                'Non-Trade Share (%)': nontrade_share
            })
                  
            fig_pie = px.pie(share_df, values=[trade_share.mean(), nontrade_share.mean()], 
                                     names=['Trade', 'Non-Trade'], title='Average Share Distribution',color=["T","NT"],color_discrete_map={"T":"green","NT":"blue"},hole=0.5)
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
    """)
from plotly.subplots import make_subplots
import matplotlib.pyplot as plt
import plotly.graph_objects as go
import plotly.express as px
import io
import requests
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.charts.legends import Legend
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
from reportlab.lib.enums import TA_CENTER
from io import BytesIO
from datetime import datetime
from reportlab.graphics import renderPDF
import random
from reportlab.lib.units import inch
from datetime import datetime
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from streamlit_lottie import st_lottie
from streamlit_option_menu import option_menu
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import Paragraph
# Sidebar navigation
def load_lottieurl(url: str):
    try:
        r = requests.get(url)
        if r.status_code != 200:
            return None
        return r.json()
    except:
        return None
def green():
 with st.sidebar:
    selected = option_menu(
        menu_title="Navigation",
        options=["Home", "Analysis", "About"],
        icons=["house", "graph-up", "info-circle"],
        menu_icon="cast",
        default_index=0,
    )
# Load Lottie animations
 lottie_analysis = load_lottieurl("https://assets4.lottiefiles.com/packages/lf20_qp1q7mct.json")
 lottie_upload = load_lottieurl("https://assets9.lottiefiles.com/packages/lf20_ABViugg1T8.json")


 def create_pdf_report(region, df, region_subset=None):
    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=letter)
    width, height = letter

    def add_page_number(canvas):
      canvas.saveState()
      canvas.setFont('Helvetica', 10)
      page_number_text = f"Page {canvas.getPageNumber()}"
      canvas.drawString(width - 100, 30, page_number_text)
      canvas.restoreState()

    # Modify the header to include region subset if provided
    def add_header(page_number):
        c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark blue color for header
        c.rect(0, height - 50, width, 50, fill=True)
        c.setFillColorRGB(1, 1, 1)  # White color for text
        c.setFont("Helvetica-Bold", 18)
        header_text = f"GYR Analysis Report: {region}"
        if region_subset:
            header_text += f" ({region_subset})"
        c.drawString(30, height - 35, header_text)

    def add_front_page():
        c.setFillColorRGB(0.4,0.5,0.3)
        c.rect(0, 0, width, height, fill=True)
        c.setFillColorRGB(1, 1, 1)
        c.setFont("Helvetica-Bold", 36)
        c.drawCentredString(width / 2, height - 200, "GYR Analysis Report")
        c.setFont("Helvetica", 24)
        report_title = f"Region: {region}"
        if region_subset:
            report_title += f" ({region_subset})"
        c.drawCentredString(width / 2, height - 250, report_title)
        c.setFont("Helvetica", 18)
        c.drawCentredString(width / 2, height - 300, f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        add_page_number(c)
        c.showPage()
    def draw_graph(fig, x, y, width, height):
        img_buffer = BytesIO()
        fig.write_image(img_buffer, format="png",scale=2)
        img_buffer.seek(0)
        img = ImageReader(img_buffer)
        c.drawImage(img, x, y, width, height)

    def draw_table(data, x, y, col_widths):
        table = Table(data, colWidths=col_widths)
        table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),  # Reduced font size
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),  # Reduced padding
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('TEXTCOLOR', (0, 1), (-1, -1), colors.black),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 6),  # Reduced font size
            ('TOPPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('BOTTOMPADDING', (0, 1), (-1, -1), 3),  # Reduced padding
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        w, h = table.wrapOn(c, width, height)
        table.drawOn(c, x, y - h)
   
    def add_tutorial_page():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Understanding the GYR Analysis")

        # Create example chart
        drawing = Drawing(400, 200)
        lc = HorizontalLineChart()
        lc.x = 40
        lc.y = 50
        lc.height = 125
        lc.width = 300
        lc.data = [
            [random.randint(2000, 3000) for _ in range(12)],  # Green
            [random.randint(1500, 2500) for _ in range(12)],  # Yellow
            [random.randint(1000, 2000) for _ in range(12)],  # Red
            [random.randint(1800, 2800) for _ in range(12)],  # Overall
            [random.randint(2200, 3200) for _ in range(12)],  # Imaginary
        ]
        lc.lines[0].strokeColor = colors.green
        lc.lines[1].strokeColor = colors.yellow
        lc.lines[2].strokeColor = colors.red
        lc.lines[3].strokeColor = colors.blue
        lc.lines[4].strokeColor = colors.purple

        # Add a legend
        legend = Legend()
        legend.alignment = 'right'
        legend.x = 330
        legend.y = 150
        legend.colorNamePairs = [
            (colors.green, 'Green EBITDA'),
            (colors.yellow, 'Yellow EBITDA'),
            (colors.red, 'Red EBITDA'),
            (colors.blue, 'Overall EBITDA'),
            (colors.purple, 'Imaginary EBITDA'),
        ]
        drawing.add(lc)
        drawing.add(legend)

        renderPDF.draw(drawing, c, inch, height - 300)

        # Key Concepts
        c.setFont("Helvetica-Bold", 18)
        c.drawString(inch, height - 350, "Key Concepts:")

        concepts = [
            ("Overall EBITDA:", "Weighted average of Green, Yellow, and Red EBITDA based on their actual quantities."),
            ("Imaginary EBITDA:", "Calculated by adjusting shares based on the following rules:"),
            ("", "• If all three (Green, Yellow, Red) are present: Green +5%, Yellow +2.5%, Red -7.5%"),
            ("", "• If only two are present: Superior one (Green in GR or GY, Yellow in YR) +5%, other -5%"),
            ("", "• If only one is present: No change"),
            ("Adjusted Shares:", "These adjustments aim to model potential improvements in product mix."),
        ]
        text_object = c.beginText(inch, height - 380)
        for title, description in concepts:
            if title:
                text_object.setFont("Helvetica-Bold", 12)
                text_object.setFillColorRGB(0.7, 0.3, 0.1)  # Reddish-brown color for concept titles
                text_object.textLine(title)
                text_object.setFont("Helvetica", 12)
                text_object.setFillColorRGB(0, 0, 0)  # Black color for descriptions
            text_object.textLine(description)
            if not title:
                text_object.textLine("")
            

        c.drawText(text_object)
        add_page_number(c)
        c.showPage()
    def add_appendix():
        c.setFont("Helvetica-Bold", 24)
        c.drawString(inch, height - inch, "Appendix")

        sections = [
            ("Graph Interpretation:", "Each line represents a different metric over time. The differences between metrics are shown below\n each month."),
            ("Tables:", "The descriptive statistics table provides a summary of the data. The monthly share distribution table\n shows the proportion of Green, Yellow, and Red products for each month."),
            ("Importance:", "These visualizations help identify trends, compare performance across product categories, and\n understand the potential impact of changing product distributions."),
        ]

        text_object = c.beginText(inch, height - 1.5*inch)
        text_object.setFont("Helvetica-Bold", 14)
        for title, content in sections:
            text_object.textLine(title)
            text_object.setFont("Helvetica", 12)
            text_object.textLines(content)
            text_object.textLine("")
            text_object.setFont("Helvetica-Bold", 14)

        c.drawText(text_object)

        # Suggestions for Improvement
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 4*inch, "Suggestions for Improvement:")

        suggestions = [
            "Increase the share of Green Region products, which typically have higher EBIDTA margins.",
            "Analyze factors contributing to higher EBIDTA in Green zone,and apply insights to Red zone.",
            "Regularly review and adjust pricing strategies to optimize EBITDA across all product categories.",
            "Invest in product innovation to expand Green and Yellow region offerings.",
        ]

        text_object = c.beginText(inch, height - 4.3*inch)
        text_object.setFont("Helvetica", 12)
        for suggestion in suggestions:
            text_object.textLine(f"• {suggestion}")

        c.drawText(text_object)

        # Limitations
        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, height - 5.2*inch, "Limitations:")

        limitations = [
            "This analysis is based on historical data and may not predict future market changes.",
            "External factors such as economic conditions are not accounted for in this report.",
            "This report analyzes the EBIDTA for GYR keeping everything else constant.",
        ]

        text_object = c.beginText(inch, height - 5.5*inch)
        text_object.setFont("Helvetica", 12)
        for limitation in limitations:
            text_object.textLine(f"• {limitation}")

        c.drawText(text_object)

        c.setFont("Helvetica", 12)
        c.drawString(inch, 2*inch, "We are currently working on including all other factors which impact the EBIDTA across GYR,")
        c.drawString(inch, 1.8*inch, "regions which will make this analysis more robust and helpful. We will also include NSR and") 
        c.drawString(inch,1.6*inch,"Contribution in our next report.")

        c.setFont("Helvetica-Bold", 14)
        c.drawString(inch, inch, "Thank You.")
        c.showPage()
    
    add_front_page()
    add_tutorial_page()
    brands = df['Brand'].unique()
    types = df['Type'].unique()
    region_subsets = df['Region subsets'].unique()

    page_number = 1
    for brand in brands:
        for product_type in types:
            for region_subset in region_subsets:
                filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                                 (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
                
                if not filtered_df.empty:
                    add_header(c)
                    cols = ['Green EBITDA', 'Yellow EBITDA', 'Red EBITDA']
                    overall_col = 'Overall EBITDA'

                    # Calculate weighted average based on actual quantities
                    total_quantity = filtered_df['Green'] + filtered_df['Yellow'] + filtered_df['Red']
                    filtered_df[overall_col] = (
                        (filtered_df['Green'] * filtered_df['Green EBITDA'] +
                         filtered_df['Yellow'] * filtered_df['Yellow EBITDA'] + 
                         filtered_df['Red'] * filtered_df['Red EBITDA']) / total_quantity
                    )

                    # Calculate current shares
                    filtered_df['Average Green Share'] = filtered_df['Green'] / total_quantity
                    filtered_df['Average Yellow Share'] = filtered_df['Yellow'] / total_quantity
                    filtered_df['Average Red Share'] = filtered_df['Red'] / total_quantity
                    
                    # Calculate Imaginary EBITDA with adjusted shares
                    def adjust_shares(row):
                        green = row['Average Green Share']
                        yellow = row['Average Yellow Share']
                        red = row['Average Red Share']
                        
                        if green == 1 or yellow == 1 or red == 1:
                            # If any share is 100%, don't change
                            return green, yellow, red
                        elif red == 0:
                            green = min(green +0.05, 1)
                            yellow = max(1-green, 0)
                        elif green == 0 and yellow == 0:
                            # If both green and yellow are absent, don't change
                            return green, yellow, red
                        elif green == 0:
                            # If green is absent, increase yellow by 5% and decrease red by 5%
                            yellow = min(yellow + 0.05, 1)
                            red = max(1 - yellow, 0)
                        elif yellow == 0:
                            # If yellow is absent, increase green by 5% and decrease red by 5%
                            green = min(green + 0.05, 1)
                            red = max(1 - green, 0)
                        else:
                            # Normal case: increase green by 5%, yellow by 2.5%, decrease red by 7.5%
                            green = min(green + 0.05, 1)
                            yellow = min(yellow + 0.025, 1 - green)
                            red = max(1 - green - yellow, 0)
                        
                        return green, yellow, red
                    filtered_df['Adjusted Green Share'], filtered_df['Adjusted Yellow Share'], filtered_df['Adjusted Red Share'] = zip(*filtered_df.apply(adjust_shares, axis=1))
                    
                    filtered_df['Imaginary EBITDA'] = (
                        filtered_df['Adjusted Green Share'] * filtered_df['Green EBITDA'] +
                        filtered_df['Adjusted Yellow Share'] * filtered_df['Yellow EBITDA'] +
                        filtered_df['Adjusted Red Share'] * filtered_df['Red EBITDA']
                    )

                    # Calculate differences
                    filtered_df['G-R Difference'] = filtered_df['Green EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['G-Y Difference'] = filtered_df['Green EBITDA'] - filtered_df['Yellow EBITDA']
                    filtered_df['Y-R Difference'] = filtered_df['Yellow EBITDA'] - filtered_df['Red EBITDA']
                    filtered_df['I-O Difference'] = filtered_df['Imaginary EBITDA'] - filtered_df[overall_col]
                    
                    # Create the plot
                    fig = go.Figure()
                    fig = make_subplots(rows=2, cols=1, row_heights=[0.58, 0.42], vertical_spacing=0.18)

                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Green EBITDA'],
                                             mode='lines+markers', name='Green EBIDTA', line=dict(color='green')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Yellow EBITDA'],
                                             mode='lines+markers', name='Yellow EBIDTA', line=dict(color='yellow')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Red EBITDA'],
                                             mode='lines+markers', name='Red EBIDTA', line=dict(color='red')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                             mode='lines+markers', name=overall_col, line=dict(color='blue', dash='dash')), row=1, col=1)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['Imaginary EBITDA'],
                                             mode='lines+markers', name='Imaginary EBIDTA',
                                             line=dict(color='purple', dash='dot')), row=1, col=1)

                    # Add I-O difference trace to the second subplot
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df['I-O Difference'],
                                             mode='lines+markers+text', name='I-O Difference',
                                             text=filtered_df['I-O Difference'].round(2),
                                             textposition='top center',textfont=dict(size=8,weight="bold"),
                                             line=dict(color='fuchsia')), row=2, col=1)

                    # Add mean line to the second subplot
                    mean_diff = filtered_df['I-O Difference'].mean()
                    if not np.isnan(mean_diff):
                        mean_diff=round(mean_diff)
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=[mean_diff] * len(filtered_df),
                                             mode='lines', name=f'Mean I-O Difference[{mean_diff}]',
                                             line=dict(color='black', dash='dash')), row=2, col=1)

                    # Customize x-axis labels for the main plot
                    x_labels = [f"{month}<br>(G-R: {g_r:.0f})<br>(G-Y: {g_y:.0f})<br>(Y-R: {y_r:.0f})" 
                                for month, g_r, g_y, y_r, i_o in 
                                zip(filtered_df['Month'], 
                                    filtered_df['G-R Difference'], 
                                    filtered_df['G-Y Difference'], 
                                    filtered_df['Y-R Difference'], 
                                    filtered_df['I-O Difference'])]

                    fig.update_layout(
                        title=f"EBITDA Analysis for {brand}({product_type}) in {region}({region_subset})",
                        legend_title='Metrics',
                        plot_bgcolor='cornsilk',
                        paper_bgcolor='lightcyan',
                        height=710,  # Increased height to accommodate the new subplot
                    )
                    fig.update_xaxes(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels, row=1, col=1)
                    fig.update_xaxes(title_text='Months', row=2, col=1)
                    fig.update_yaxes(title_text='EBITDA(Rs./MT)', row=1, col=1)
                    fig.update_yaxes(title_text='I-O Difference(Rs./MT)', row=2, col=1)
                    # Add new page if needed
                    #if page_number > 1:
                        #c.showPage()
                    # Draw the graph
                    draw_graph(fig, 50, height - 410, 500, 350)

                    # Add descriptive statistics
                    c.setFillColorRGB(0.2, 0.2, 0.7)  # Dark grey color for headers
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 425, "Descriptive Statistics")
                    
                    desc_stats = filtered_df[['Green','Yellow','Red']+cols + [overall_col, 'Imaginary EBITDA']].describe().reset_index()
                    desc_stats = desc_stats[desc_stats['index'] != 'count'].round(2)  # Remove 'count' row
                    table_data = [['Metric'] + list(desc_stats.columns[1:])] + desc_stats.values.tolist()
                    draw_table(table_data, 50, height - 435, [40,40,40,40] + [75] * (len(desc_stats.columns) - 4))  # Reduced column widths
                    c.setFont("Helvetica-Bold", 10)  # Reduced font size
                    c.drawString(50, height - 600, "Average Share Distribution")
                
                    # Create pie chart with correct colors
                    average_shares = filtered_df[['Average Green Share', 'Average Yellow Share', 'Average Red Share']].mean()
                    share_fig = px.pie(
                       values=average_shares.values,
                       names=average_shares.index,
                       color=average_shares.index,
                       color_discrete_map={'Average Green Share': 'green', 'Average Yellow Share': 'yellow', 'Average Red Share': 'red'},
                       title="",hole=0.3)
                    share_fig.update_layout(width=475, height=475, margin=dict(l=0, r=0, t=0, b=0))  # Reduced size
                    
                    draw_graph(share_fig, 80, height - 810, 200, 200)  # Adjusted position and size
                    c.setFont("Helvetica-Bold", 10)
                    c.drawString(330, height - 600, "Monthly Share Distribution")
                    share_data = [['Month', 'Green', 'Yellow', 'Red']]
                    for _, row in filtered_df[['Month', 'Green', 'Yellow', 'Red', 'Average Green Share', 'Average Yellow Share', 'Average Red Share']].iterrows():
                        share_data.append([
                            row['Month'],
                            f"{row['Green']:.0f} ({row['Average Green Share']:.2%})",
                            f"{row['Yellow']:.0f} ({row['Average Yellow Share']:.2%})",
                            f"{row['Red']:.0f} ({row['Average Red Share']:.2%})"
                        ])
                    draw_table(share_data, 330, height - 620, [40, 60, 60, 60])
                    add_page_number(c)
                    c.showPage()
    for i in range(c.getPageNumber()):
        c.setPageSize((width, height))
        add_page_number(c)         
    add_appendix()
    c.save()
    buffer.seek(0)
    return buffer


 if selected == "Home":
    st.title("📊 Advanced GYR Analysis")
    st.markdown("Welcome to our advanced data analysis platform. Upload your Excel file to get started with interactive visualizations and insights.")
    
    st.markdown("<div class='upload-section'>", unsafe_allow_html=True)
    col1, col2 = st.columns([2, 1])
    with col1:
        uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx",key="gyruploader")
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
    st.title("📈 GYR Dashboard")
    
    if 'uploaded_file' not in st.session_state or st.session_state.uploaded_file is None:
        st.warning("Please upload an Excel file on the Home page to begin the analysis.")
    else:
        df = pd.read_excel(st.session_state.uploaded_file)
        st.markdown("<div class='analysis-section'>", unsafe_allow_html=True)
        
        if lottie_analysis:
            st_lottie(lottie_analysis, height=200, key="analysis")
        else:
            st.image("https://cdn-icons-png.flaticon.com/512/2756/2756778.png", width=200)
        st.sidebar.header("Filter Options")
        region = st.sidebar.selectbox("Select Region", options=df['Region'].unique(), key="region_select")

        # Add download options for report
        st.sidebar.subheader(f"Download Report for {region}")
        download_choice = st.sidebar.radio(
            "Choose report type:",
            ('Full Region', 'Region Subset')
        )
        
        if download_choice == 'Full Region':
            if st.sidebar.button(f"Download Full Report for {region}"):
                pdf_buffer = create_pdf_report(region, df)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}.pdf">Download Full Region PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)
        else:
            region_subsets = df[df['Region'] == region]['Region subsets'].unique()
            selected_subset = st.sidebar.selectbox("Select Region Subset", options=region_subsets)
            if st.sidebar.button(f"Download Report for {region} - {selected_subset}"):
                # Filter the dataframe for the selected region and subset
                subset_df = df[(df['Region'] == region) & (df['Region subsets'] == selected_subset)]
                pdf_buffer = create_pdf_report(region, subset_df, selected_subset)
                pdf_bytes = pdf_buffer.getvalue()
                b64 = base64.b64encode(pdf_bytes).decode()
                href = f'<a href="data:application/pdf;base64,{b64}" download="GYR_Analysis_Report_{region}_{selected_subset}.pdf">Download Region Subset PDF Report</a>'
                st.sidebar.markdown(href, unsafe_allow_html=True)

        # Add unique keys to each selectbox
        brand = st.sidebar.selectbox("Select Brand", options=df[df['Region']==region]['Brand'].unique(), key="brand_select")
        product_type = st.sidebar.selectbox("Select Type", options=df[df['Region']==region]['Type'].unique(), key="type_select")
        region_subset = st.sidebar.selectbox("Select Region Subset", options=df[df['Region']==region]['Region subsets'].unique(), key="region_subset_select")

        
        # Analysis type selection using radio buttons
        st.sidebar.header("Analysis on")
        analysis_options = ["NSR Analysis", "Contribution Analysis", "EBITDA Analysis"]
        
        # Use session state to store the selected analysis type
        if 'analysis_type' not in st.session_state:
            st.session_state.analysis_type = "EBITDA Analysis"
        
        analysis_type = st.sidebar.radio("Select Analysis Type", analysis_options, index=analysis_options.index(st.session_state.analysis_type), key="analysis_type_radio")
        
        # Update session state
        st.session_state.analysis_type = analysis_type

        green_share = st.sidebar.slider("Adjust Green Share (%)", 0, 99, 50, key="green_share_slider")
        yellow_share = st.sidebar.slider("Adjust Yellow Share (%)", 0, 100-green_share, 0, key="yellow_share_slider")
        red_share = 100 - green_share - yellow_share
        st.sidebar.text(f"Red Share: {red_share}%")
        # Filter the dataframe
        filtered_df = df[(df['Region'] == region) & (df['Brand'] == brand) &
                         (df['Type'] == product_type) & (df['Region subsets'] == region_subset)].copy()
        
        if not filtered_df.empty:
            if analysis_type == 'NSR Analysis':
                cols = ['Green NSR', 'Yellow NSR', 'Red NSR']
                overall_col = 'Overall NSR'
            elif analysis_type == 'Contribution Analysis':
                cols = ['Green Contribution', 'Yellow Contribution','Red Contribution']
                overall_col = 'Overall Contribution'
            elif analysis_type == 'EBITDA Analysis':
                cols = ['Green EBITDA', 'Yellow EBITDA','Red EBITDA']
                overall_col = 'Overall EBITDA'
            
            # Calculate weighted average based on actual quantities
            filtered_df[overall_col] = (filtered_df['Green'] * filtered_df[cols[0]] +
                                        filtered_df['Yellow'] * filtered_df[cols[1]] + filtered_df['Red']*filtered_df[cols[2]]) / (
                                            filtered_df['Green'] + filtered_df['Yellow']+filtered_df['Red'])
            
            # Calculate imaginary overall based on slider
            imaginary_col = f'Imaginary {overall_col}'
            filtered_df[imaginary_col] = ((1 - (green_share+yellow_share)/100) * filtered_df[cols[2]] +
                                          (green_share/100) * filtered_df[cols[0]] + (yellow_share/100) * filtered_df[cols[1]])
            
            # Calculate difference between Premium and Normal
            filtered_df['G-Y Difference'] = filtered_df[cols[0]] - filtered_df[cols[1]]
            filtered_df['G-R Difference'] = filtered_df[cols[0]] - filtered_df[cols[2]]
            filtered_df['Y-R Difference'] = filtered_df[cols[1]] - filtered_df[cols[2]]
            
            # Calculate difference between Imaginary and Overall
            filtered_df['Imaginary vs Overall Difference'] = filtered_df[imaginary_col] - filtered_df[overall_col]
            
            # Create the plot
            fig = go.Figure()
            
            
            if cols[0] in cols:
                  fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[0]],
                                         mode='lines+markers', name=cols[0],line_color="green"))
            if cols[1] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[1]],
                                         mode='lines+markers', name=cols[1],line_color="yellow"))
            if cols[2] in cols:
                    fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[cols[2]],
                                         mode='lines+markers', name=cols[2],line_color="red"))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[overall_col],
                                     mode='lines+markers', name=overall_col, line=dict(dash='dash')))
            
            fig.add_trace(go.Scatter(x=filtered_df['Month'], y=filtered_df[imaginary_col],
                                     mode='lines+markers', name=f'Imaginary {overall_col} ({green_share}% Green & {yellow_share}% Yellow)',
                                     line=dict(color='brown', dash='dot')))
            
            # Customize x-axis labels to include the differences
            x_labels = [f"{month}<br>(G-Y: {diff:.2f})<br>(G-R: {i_diff:.2f})<br>(Y-R: {j_diff:.2f})<br>(I-O: {k_diff:.2f})" for month, diff, i_diff, j_diff, k_diff in 
                        zip(filtered_df['Month'], filtered_df['G-Y Difference'], filtered_df['G-R Difference'], filtered_df['Y-R Difference'], filtered_df['Imaginary vs Overall Difference'])]
            
            fig.update_layout(
                title=analysis_type,
                xaxis_title='Month (G-Y: Green - Red,G-R: Green - Red,Y-R: Yellow - Red, I-O: Imaginary - Overall)',
                yaxis_title='Value',
                legend_title='Metrics',
                hovermode="x unified",
                xaxis=dict(tickmode='array', tickvals=list(range(len(x_labels))), ticktext=x_labels)
            )
            
            st.plotly_chart(fig, use_container_width=True)
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
                    'Red Share (%)': red_share,
                    'Green Quantity': filtered_df['Green'],
                    'Yellow Quantity': filtered_df['Yellow'],
                    'Red Quantity': filtered_df['Red']})

# Function to create pie chart
    def create_pie_chart(row):
     values = [row['Green Quantity'], row['Yellow Quantity'], row['Red Quantity']]
     labels = ['Green', 'Yellow', 'Red']
     colors = ['green', 'yellow', 'red']
    
     fig = px.pie(
        values=values,
        names=labels,
        title=f"Product Distribution for {row['Month']}",
        color=labels,
        color_discrete_map=dict(zip(labels, colors)),
        hole=0.3
    )
    
     total_quantity = sum(values)
     fig.add_annotation(
        text=f'Total Quantity: {total_quantity}',
        x=0.5, y=-0.15,
        showarrow=False
    )
    
    return fig

# Display interactive dataframe
    st.markdown("Click on a row to see detailed distribution")
    selected_indices = st.data_editor(
    share_df.set_index('Month')[['Green Share (%)', 'Yellow Share (%)', 'Red Share (%)']],
    hide_index=False,
    use_container_width=True,
    key="interactive_share_df",
    disabled=True)

# Display pie chart for selected row
    if selected_indices:
              selected_month = list(selected_indices.keys())[0]
              selected_row = share_df[share_df['Month'] == selected_month].iloc[0]
              st.plotly_chart(create_pie_chart(selected_row), use_container_width=True)

# Display average distribution pie chart
    fig_avg_pie = px.pie(
                     values=[green_share.mean(), yellow_share.mean(), red_share.mean()],
                     names=['Green', 'Yellow', 'Red'],
                     title='Average Share Distribution',
                     color=['Green', 'Yellow', 'Red'],
                     color_discrete_map={"Green": "green", "Yellow": "yellow", "Red": "red"},
                     hole=0.3)
    st.plotly_chart(fig_avg_pie, use_container_width=True)   
    fig_pie = px.pie(share_df, values=[green_share.mean(), yellow_share.mean(), red_share.mean()], 
                                     names=['Green', 'Yellow', 'Red'], title='Average Share Distribution',color=["G","Y","R"],color_discrete_map={"G":"green","Y":"yellow","R":"red"},hole=0.5)
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
    """)

def load_visit_data():
    try:
        with open('visit_data.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {'total_visits': 0, 'daily_visits': {}}

def save_visit_data(data):
    with open('visit_data.json', 'w') as f:
        json.dump(data, f)
def update_visit_count():
    visit_data = load_visit_data()
    today = datetime.now().strftime('%Y-%m-%d')
    
    visit_data['total_visits'] += 1
    visit_data['daily_visits'][today] = visit_data['daily_visits'].get(today, 0) + 1
    
    save_visit_data(visit_data)
    return visit_data['total_visits'], visit_data['daily_visits'][today]
def load_visit_data():
    try:
        with open('visit_data.json', 'r') as f:
            return json.load(f)
    except FileNotFoundError:
        return {'total_visits': 0, 'daily_visits': {}}

def save_visit_data(data):
    with open('visit_data.json', 'w') as f:
        json.dump(data, f)

def update_visit_count():
    visit_data = load_visit_data()
    today = datetime.now().strftime('%Y-%m-%d')
    
    visit_data['total_visits'] += 1
    visit_data['daily_visits'][today] = visit_data['daily_visits'].get(today, 0) + 1
    
    save_visit_data(visit_data)
    return visit_data['total_visits'], visit_data['daily_visits'][today]

def main():
    # Set background image (uncomment and replace 'background.png' with your image file)
    # set_png_as_page_bg('background.png')

    # Custom CSS for the sidebar and main content
    st.markdown("""
    <style>
    .sidebar .sidebar-content {
        background-image: linear-gradient(180deg, #2e7bcf 25%, #4527A0 100%);
        color: white;
    }
    .sidebar-text {
        color: white !important;
    }
    .stButton>button {
        width: 100%;
        border-radius: 20px;
        background-color: #4CAF50;
        color: white;
        border: none;
        padding: 10px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #45a049;
        box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1);
    }
    .stProgress .st-bo {
        background-color: #4CAF50;
    }
    .stProgress .st-bp {
        background-color: #E0E0E0;
    }
    .settings-container {
        background-color: rgba(255, 255, 255, 0.1);
        backdrop-filter: blur(10px);
        padding: 20px;
        border-radius: 10px;
        margin-top: 20px;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .visit-counter {
        background-color: rgba(255, 228, 225, 0.7);
        border-radius: 10px;
        padding: 15px;
        margin-top: 20px;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1);
    }
    .visit-counter h3 {
        color: #333;
        font-size: 18px;
        margin-bottom: 10px;
    }
    .visit-counter p {
        color: #8B4513;
        font-size: 14px;
        margin: 5px 0;
    }
    .user-info {
        background-color: rgba(255, 255, 255, 0.1);
        border-radius: 10px;
        padding: 10px;
        margin-bottom: 20px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.sidebar.title("Analytics Dashboard")
    
    # User info with avatar
    if 'username' not in st.session_state:
        st.session_state.username = "Guest"
    

    
    st.sidebar.markdown(f"""
    <div class="user-info">
        <i class="fas fa-user"></i> Logged in as: {st.session_state.username}
    </div>
    """, unsafe_allow_html=True)

    # Main menu with icons and hover effects
    with st.sidebar:
        selected = option_menu(
            menu_title="Main Menu",
            options=[
                "Home", 
                "Data Management", 
                "Analysis Dashboards", 
                "Predictions", 
                "Settings"
            ],
            icons=[
                "house-fill", 
                "database-fill-gear", 
                "graph-up-arrow", 
                "lightbulb-fill", 
                "gear-fill"
            ],
            menu_icon="cast",
            default_index=0,
            styles={
                "container": {"padding": "0!important", "background-color": "transparent"},
                "icon": {"color": "orange", "font-size": "20px"}, 
                "nav-link": {"font-size": "16px", "text-align": "left", "margin":"0px", "--hover-color": "#eee"},
                "nav-link-selected": {"background-color": "rgba(255, 255, 255, 0.2)"},
            }
        )

    # Submenu based on main selection
    if selected == "Home":
        Home()
    elif selected == "Data Management":
        data_management_menu = option_menu(
            menu_title="Data Management",
            options=["Editor", "File Manager"],
            icons=["pencil-square", "folder"],
            orientation="horizontal",
        )
        if data_management_menu == "Editor":
            excel_editor_and_analyzer()
        elif data_management_menu == "File Manager":
            folder_menu()
    elif selected == "Analysis Dashboards":
        analysis_menu = option_menu(
            menu_title="Analysis Dashboards",
            options=["WSP Analysis", "Sales Dashboard", "Product-Mix", "Segment-Mix","Geo-Mix"],
            icons=["clipboard-data", "cash", "arrow-up-right", "shuffle", "globe"],
            orientation="horizontal",
        )
        if analysis_menu == "WSP Analysis":
            wsp_analysis_dashboard()
        elif analysis_menu == "Sales Dashboard":
            sales_dashboard()
        elif analysis_menu == "Product-Mix":
            normal()
        elif analysis_menu == "Segment-Mix":
            trade()
        elif analysis_menu == "Geo-Mix":
            green()
    elif selected == "Predictions":
        prediction_menu = option_menu(
            menu_title="Predictions",
            options=["Descriptive Statistics", "Sales Prediction"],
            icons=["bar-chart", "graph-up-arrow"],
            orientation="horizontal",
        )
        if prediction_menu == "Descriptive Statistics":
            descriptive_statistics_and_prediction()
        elif prediction_menu == "Sales Prediction":
            sales_prediction_app()
    elif selected == "Settings":
        st.title("Settings")
        st.markdown('<div class="settings-container">', unsafe_allow_html=True)
        
        # User Settings
        st.subheader("User Settings")
        username = st.text_input("Username", value=st.session_state.username)
        email = st.text_input("Email", value="johndoe@example.com")
        if st.button("Update Profile"):
            st.session_state.username = username
            st.success("Profile updated successfully!")
        
        # Appearance Settings
        st.subheader("Appearance")
        theme = st.selectbox("Theme", ["Light", "Dark", "System Default"])
        chart_color = st.color_picker("Default Chart Color", "#2e7bcf")
        
        # Notification Settings
        st.subheader("Notifications")
        email_notifications = st.checkbox("Receive Email Notifications", value=True)
        notification_frequency = st.select_slider("Notification Frequency", options=["Daily", "Weekly", "Monthly"])
        
        # Save Settings Button
        if st.button("Save Settings"):
            st.success("Settings saved successfully!")
        
        st.markdown('</div>', unsafe_allow_html=True)

    # Enhanced feedback section
    st.sidebar.markdown("---")
    st.sidebar.subheader("📢 Feedback")
    feedback = st.sidebar.text_area("Share your thoughts:")
    if st.sidebar.button("Submit Feedback", key="submit_feedback"):
        # Here you would typically send this feedback to a database or email
        st.sidebar.success("Thank you for your valuable feedback!")
    

    # Display visit counter with animations
    total_visits, daily_visits = update_visit_count()
    st.sidebar.markdown(f"""
    <div class="visit-counter">
        <h3>📊 Visit Statistics</h3>
        <p>Total Visits: <span class="count">{total_visits}</span></p>
        <p>Visits Today: <span class="count">{daily_visits}</span></p>
    </div>
    <script>
        const countElements = document.querySelectorAll('.count');
        countElements.forEach(element => {{
            const target = parseInt(element.innerText);
            let count = 0;
            const timer = setInterval(() => {{
                element.innerText = count;
                if (count === target) {{
                    clearInterval(timer);
                }}
                count++;
            }}, 20);
        }});
    </script>
    """, unsafe_allow_html=True)
if __name__ == "__main__":
    main()
