import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split, RandomizedSearchCV
from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from xgboost import XGBRegressor
from lightgbm import LGBMRegressor
from sklearn.linear_model import ElasticNet
from sklearn.svm import SVR
from scipy.stats import randint, uniform
import base64
from io import BytesIO

# Load and preprocess data
@st.cache_data
def load_data(uploaded_file):
    df = pd.read_excel(uploaded_file)
    return df

def preprocess_data(df):
    df['YoY_Growth'] = (df['Monthly Achievement(Aug)'] - df['Total Aug 2023']) / df['Total Aug 2023']
    
    months = ['Apr', 'May', 'June', 'July', 'Aug']
    for i in range(1, len(months)):
        df[f'Growth_{months[i-1]}_{months[i]}'] = (df[f'Monthly Achievement({months[i]})'] - df[f'Monthly Achievement({months[i-1]})']) / df[f'Monthly Achievement({months[i-1]})']
    
    growth_cols = [f'Growth_{months[i-1]}_{months[i]}' for i in range(1, len(months))]
    df['Avg_Monthly_Growth'] = df[growth_cols].mean(axis=1)
    
    for month in months:
        df[f'Achievement_Rate_{month}'] = df[f'Monthly Achievement({month})'] / df[f'Month Tgt ({month})']
    
    achievement_rate_cols = [f'Achievement_Rate_{month}' for month in months]
    df['Avg_Achievement_Rate'] = df[achievement_rate_cols].mean(axis=1)
    
    df['Seasonal_Index'] = df['Monthly Achievement(Aug)'] / df['Total Sep 2023']
    df['Aug_YoY_Diff'] = df['Monthly Achievement(Aug)'] - df['Total Aug 2023']
    df['Sep_Aug_Target_Ratio'] = df['Month Tgt (Sep)'] / df['Monthly Achievement(Aug)']
    
    return df

def create_features_target(df):
    features = [
        'Month Tgt (Sep)', 'Monthly Achievement(Aug)', 'Total Sep 2023', 'Total Aug 2023',
        'YoY_Growth', 'Avg_Monthly_Growth', 'Avg_Achievement_Rate', 'Seasonal_Index',
        'Aug_YoY_Diff', 'Sep_Aug_Target_Ratio'
    ] + [f'Month Tgt ({month})' for month in ['Apr', 'May', 'June', 'July', 'Aug']]
    
    X = df[features]
    y = df['Monthly Achievement(Aug)']
    return X, y

# Model training and prediction
@st.cache_resource
def train_ensemble_model(X, y):
    models = {
        'rf': RandomForestRegressor(random_state=42),
        'gb': GradientBoostingRegressor(random_state=42),
        'xgb': XGBRegressor(random_state=42),
        'lgbm': LGBMRegressor(random_state=42),
        'elastic': ElasticNet(random_state=42),
        'svr': SVR()
    }
    
    param_spaces = {
        'rf': {'n_estimators': randint(100, 1000), 'max_depth': randint(5, 30)},
        'gb': {'n_estimators': randint(100, 1000), 'learning_rate': uniform(0.01, 0.3)},
        'xgb': {'n_estimators': randint(100, 1000), 'learning_rate': uniform(0.01, 0.3)},
        'lgbm': {'n_estimators': randint(100, 1000), 'learning_rate': uniform(0.01, 0.3)},
        'elastic': {'alpha': uniform(0, 1), 'l1_ratio': uniform(0, 1)},
        'svr': {'C': uniform(0.1, 10), 'epsilon': uniform(0.01, 0.1)}
    }
    
    best_models = {}
    for name, model in models.items():
        random_search = RandomizedSearchCV(model, param_spaces[name], n_iter=20, cv=5, n_jobs=-1, random_state=42)
        random_search.fit(X, y)
        best_models[name] = random_search.best_estimator_
    
    return best_models

def ensemble_predict(models, X):
    predictions = np.column_stack([model.predict(X) for model in models.values()])
    return np.mean(predictions, axis=1)

def calculate_metrics(y_true, y_pred):
    mse = mean_squared_error(y_true, y_pred)
    rmse = np.sqrt(mse)
    mae = mean_absolute_error(y_true, y_pred)
    r2 = r2_score(y_true, y_pred)
    return rmse, mae, r2

def predict_sales(df, region, brand):
    df_filtered = df[(df['Zone'] == region) & (df['Brand'] == brand)].copy()
    
    if len(df_filtered) == 0:
        return None, None, None, None
    
    df_processed = preprocess_data(df_filtered)
    X, y = create_features_target(df_processed)
    
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    
    scaler = StandardScaler()
    X_train_scaled = scaler.fit_transform(X_train)
    X_test_scaled = scaler.transform(X_test)
    
    best_models = train_ensemble_model(X_train_scaled, y_train)
    
    y_pred = ensemble_predict(best_models, X_test_scaled)
    
    rmse, mae, r2 = calculate_metrics(y_test, y_pred)
    
    sept_features = X.iloc[-1].values.reshape(1, -1)
    sept_features_scaled = scaler.transform(sept_features)
    
    model_predictions = []
    for model in best_models.values():
        try:
            pred = model.predict(sept_features_scaled)[0]
            model_predictions.append(pred)
        except Exception as e:
            st.error(f"Error in model prediction: {str(e)}")
    
    if not model_predictions:
        return None, None, None, None
    
    sept_prediction = np.mean(model_predictions)
    
    ci_lower, ci_upper = np.percentile(model_predictions, [2.5, 97.5])
    
    n_iterations = 1000
    bootstrap_predictions = []
    for _ in range(n_iterations):
        bootstrap_sample = np.random.choice(model_predictions, size=len(model_predictions), replace=True)
        bootstrap_predictions.append(np.mean(bootstrap_sample))
    
    pi_lower, pi_upper = np.percentile(bootstrap_predictions, [2.5, 97.5])
    
    return sept_prediction, (ci_lower, ci_upper), (pi_lower, pi_upper), rmse

# Visualization
def create_visualization(df, region, brand, prediction, ci, pi):
    fig, ax = plt.subplots(figsize=(12, 6))
    
    months = ['Apr', 'May', 'June', 'July', 'Aug', 'Sep']
    achievements = [df[f'Monthly Achievement({m})'].iloc[-1] for m in months[:-1]] + [prediction]
    targets = [df[f'Month Tgt ({m})'].iloc[-1] for m in months]
    
    ax.bar(months, targets, alpha=0.5, label='Target')
    ax.bar(months, achievements, alpha=0.7, label='Achievement')
    
    ax.set_title(f'Monthly Targets and Achievements for {region} - {brand}')
    ax.set_xlabel('Month')
    ax.set_ylabel('Sales')
    ax.legend()
    
    ax.errorbar('Sep', prediction, yerr=[[prediction-ci[0]], [ci[1]-prediction]], 
                fmt='o', color='r', capsize=5, label='95% CI')
    ax.errorbar('Sep', prediction, yerr=[[prediction-pi[0]], [pi[1]-prediction]], 
                fmt='o', color='g', capsize=5, label='95% PI')
    
    ax.legend()
    
    return fig

# Streamlit app
def main():
    st.set_page_config(page_title="Sales Prediction App", page_icon="📊", layout="wide")
    
    st.title("📊 Sales Prediction App")
    
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        df = load_data(uploaded_file)
        
        st.sidebar.header("Filters")
        region = st.sidebar.selectbox("Select Region", df['Zone'].unique())
        brand = st.sidebar.selectbox("Select Brand", df['Brand'].unique())
        
        if st.sidebar.button("Generate Prediction"):
            with st.spinner("Generating prediction..."):
                prediction, ci, pi, rmse = predict_sales(df, region, brand)
            
            if prediction is not None:
                st.success("Prediction generated successfully!")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.subheader("Prediction Results")
                    st.write(f"Predicted September 2024 sales: {prediction:.2f}")
                    st.write(f"95% Confidence Interval: ({ci[0]:.2f}, {ci[1]:.2f})")
                    st.write(f"95% Prediction Interval: ({pi[0]:.2f}, {pi[1]:.2f})")
                    st.write(f"Root Mean Square Error (RMSE): {rmse:.2f}")
                
                with col2:
                    fig = create_visualization(df, region, brand, prediction, ci, pi)
                    st.pyplot(fig)
                
                st.subheader("Interpretation")
                relative_ci_width = (ci[1] - ci[0]) / prediction * 100
                relative_pi_width = (pi[1] - pi[0]) / prediction * 100
                st.write(f"Relative Confidence Interval width: {relative_ci_width:.2f}%")
                st.write(f"Relative Prediction Interval width: {relative_pi_width:.2f}%")
                
                if relative_ci_width < 10:
                    st.write("The model's confidence interval is narrow, indicating high precision in the estimate.")
                elif relative_ci_width < 20:
                    st.write("The model's confidence interval is moderately narrow, indicating good precision in the estimate.")
                else:
                    st.write("The model's confidence interval is wide, indicating uncertainty in the estimate.")
                
                mean_sales = df['Monthly Achievement(Aug)'].mean()
                if rmse < 0.1 * mean_sales:
                    st.write("The RMSE is low relative to the mean sales, indicating good model performance.")
                elif rmse < 0.2 * mean_sales:
                    st.write("The RMSE is moderate relative to the mean sales, indicating acceptable model performance.")
                else:
                    st.write("The RMSE is high relative to the mean sales, indicating that the model may need improvement.")
            else:
                st.error("Unable to generate prediction. Please check your data and selected filters.")
        
        st.sidebar.header("Prediction Techniques")
        st.sidebar.write("""
        1. Ensemble Learning: Combines predictions from multiple models to reduce bias and variance.
           $f_{ensemble}(x) = \frac{1}{M} \sum_{i=1}^M f_i(x)$

        2. Feature Engineering: Creates new features to capture complex patterns in the data.
           e.g., YoY Growth = $\frac{Aug_{2024} - Aug_{2023}}{Aug_{2023}}$

        3. Hyperparameter Tuning: Uses RandomizedSearchCV to optimize model parameters.
           $\theta^* = \arg\min_{\theta} \sum_{i=1}^n L(y_i, f(x_i; \theta))$
        """)

if __name__ == "__main__":
    main()
