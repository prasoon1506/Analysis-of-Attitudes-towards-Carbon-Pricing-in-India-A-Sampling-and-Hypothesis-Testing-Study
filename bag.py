import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import seaborn as sns
import numpy as np
from datetime import datetime
from scipy import stats
from statsmodels.tsa.seasonal import seasonal_decompose

# Helper Functions
def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def calculate_plant_statistics(df, plant_name):
    """Calculate comprehensive plant-level statistics"""
    plant_data = df[df['Cement Plant Sname'] == plant_name]
    month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
    
    # Calculate total usage per bag
    bag_totals = {}
    for _, row in plant_data.iterrows():
        bag_totals[row['MAKTX']] = row[month_columns].sum()
    
    # Sort bags by total usage
    sorted_bags = dict(sorted(bag_totals.items(), key=lambda x: x[1], reverse=True))
    
    # Calculate bag mix percentages
    total_plant_usage = sum(bag_totals.values())
    bag_mix = {bag: (usage/total_plant_usage)*100 for bag, usage in sorted_bags.items()}
    
    # Calculate month-over-month growth rates
    monthly_totals = []
    for month in month_columns:
        monthly_totals.append(plant_data[month].sum())
    
    mom_growth = pd.Series(monthly_totals).pct_change().mean() * 100
    
    # Calculate volatility (coefficient of variation)
    volatility = np.std(monthly_totals) / np.mean(monthly_totals) * 100
    
    # Calculate capacity utilization
    max_monthly = max(monthly_totals)
    theoretical_capacity = max_monthly * 1.2
    avg_utilization = (np.mean(monthly_totals) / theoretical_capacity) * 100
    
    # Calculate seasonality strength
    monthly_series = pd.Series(monthly_totals)
    if len(monthly_series) >= 12:
        seasonal_strength = calculate_seasonality_strength(monthly_series)
    else:
        seasonal_strength = None
    
    return {
        'Total Plant Usage': sum(monthly_totals),
        'Average Monthly Usage': np.mean(monthly_totals),
        'Peak Monthly Usage': max_monthly,
        'Lowest Monthly Usage': min(monthly_totals),
        'Number of Bag Types': len(bag_totals),
        'Top Bags': dict(list(sorted_bags.items())[:3]),
        'Bag Mix': bag_mix,
        'MoM Growth Rate': mom_growth,
        'Usage Volatility': volatility,
        'Capacity Utilization': avg_utilization,
        'Seasonality Strength': seasonal_strength
    }

def calculate_seasonality_strength(series):
    """Calculate seasonality strength using STL decomposition"""
    # Perform decomposition
    decomposition = seasonal_decompose(series, period=12, extrapolate_trend='freq')
    
    # Calculate strength of seasonality
    seasonal_strength = np.std(decomposition.seasonal) / np.std(series) * 100
    return seasonal_strength

def analyze_bag_patterns(df, plant_name):
    """Analyze usage patterns and relationships between bags"""
    plant_data = df[df['Cement Plant Sname'] == plant_name]
    month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
    
    # Create usage matrix
    usage_matrix = plant_data[month_columns].T
    usage_matrix.columns = plant_data['MAKTX']
    
    # Calculate correlations
    correlations = usage_matrix.corr()
    
    # Identify complementary and substitute products
    high_positive_corr = correlations.unstack()
    high_positive_corr = high_positive_corr[high_positive_corr != 1.0]
    high_positive_corr = high_positive_corr[high_positive_corr > 0.7]
    
    high_negative_corr = correlations.unstack()
    high_negative_corr = high_negative_corr[high_negative_corr < -0.3]
    
    # Calculate stability scores
    stability_scores = {}
    for bag in usage_matrix.columns:
        cv = usage_matrix[bag].std() / usage_matrix[bag].mean()
        stability_scores[bag] = (1 - cv) * 100
    
    return {
        'correlations': correlations,
        'complementary_products': high_positive_corr,
        'substitute_products': high_negative_corr,
        'stability_scores': stability_scores
    }

def calculate_efficiency_metrics(df, plant_name):
    """Calculate efficiency metrics for the plant"""
    plant_data = df[df['Cement Plant Sname'] == plant_name]
    month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
    
    monthly_totals = []
    monthly_mix_entropy = []
    
    for month in month_columns:
        month_usage = plant_data[month]
        total = month_usage.sum()
        monthly_totals.append(total)
        
        # Calculate mix entropy
        if total > 0:
            proportions = month_usage / total
            entropy = -np.sum(proportions * np.log2(proportions + 1e-10))
            monthly_mix_entropy.append(entropy)
        else:
            monthly_mix_entropy.append(0)
    
    # Calculate efficiency metrics
    output_stability = 100 - (np.std(monthly_totals) / np.mean(monthly_totals) * 100)
    mix_complexity = np.mean(monthly_mix_entropy)
    
    # Calculate capacity consistency
    max_capacity = max(monthly_totals)
    target_capacity = max_capacity * 0.8
    deviations = [abs(x - target_capacity) / target_capacity for x in monthly_totals]
    capacity_consistency = 100 - (np.mean(deviations) * 100)
    
    return {
        'output_stability': output_stability,
        'mix_complexity': mix_complexity,
        'capacity_consistency': capacity_consistency,
        'monthly_mix_entropy': dict(zip(month_columns, monthly_mix_entropy))
    }

def create_forecast_plot(df, plant_name, selected_bag):
    """Create forecast plot with actual and projected values"""
    selected_data = df[(df['Cement Plant Sname'] == plant_name) & 
                      (df['MAKTX'] == selected_bag)]
    
    month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
    usage_data = []
    
    for month in month_columns:
        date = pd.to_datetime(month)
        usage = selected_data[month].iloc[0]
        usage_data.append({
            'Date': date,
            'Usage': usage
        })
    
    usage_df = pd.DataFrame(usage_data)
    usage_df = usage_df.sort_values('Date')
    usage_df['Month'] = usage_df['Date'].apply(format_date_for_display)
    
    # Create figure
    fig = go.Figure()
    
    # Add actual usage line
    fig.add_trace(go.Scatter(
        x=usage_df['Month'],
        y=usage_df['Usage'],
        name='Actual Usage',
        line=dict(color='#2E86C1', width=3),
        mode='lines+markers'
    ))
    
    # Update layout
    fig.update_layout(
        title=f'Usage Trend for {selected_bag}',
        xaxis_title='Month',
        yaxis_title='Usage',
        hovermode='x unified',
        showlegend=True
    )
    
    return fig

def main():
    # Page configuration
    st.set_page_config(
        page_title="Advanced Cement Plant Analytics",
        layout='wide',
        initial_sidebar_state='expanded'
    )

    # Custom CSS
    st.markdown("""
        <style>
        .main {
            padding: 2rem;
        }
        .stTitle {
            font-size: 2.5rem !important;
            padding-bottom: 2rem;
        }
        .stats-card {
            background-color: #f8f9fa;
            padding: 1.5rem;
            border-radius: 10px;
            box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
        }
        .announcement {
            background-color: #e3f2fd;
            padding: 1rem;
            border-radius: 10px;
            border-left: 5px solid #1976d2;
            margin: 1rem 0;
        }
        div[data-testid="stMetricValue"] {
            font-size: 1.8rem;
        }
        </style>
    """, unsafe_allow_html=True)

    # Title
    st.title("🏭 Advanced Cement Plant Analytics Platform")
    
    # Announcement
    st.markdown("""
        <div class='announcement'>
            <h3>🤖 Advanced Analytics Suite</h3>
            <p>Features included:
               • Comprehensive Plant Statistics
               • Bag Pattern Analysis
               • Efficiency Metrics
               • Usage Forecasting
               • Correlation Analysis</p>
        </div>
    """, unsafe_allow_html=True)
    
    # File uploader
    with st.sidebar:
        st.header("📁 Data Input")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            # Load and process data
            df = pd.read_excel(uploaded_file)
            df = df.iloc[:, 1:]  # Remove first column
            
            # Sidebar controls
            with st.sidebar:
                st.header("🎯 Analysis Controls")
                
                # Plant selection
                unique_plants = sorted(df['Cement Plant Sname'].unique())
                selected_plant = st.selectbox('Select Plant:', unique_plants)
                
                # Analysis type
                analysis_type = st.radio(
                    "Analysis Focus:",
                    ["Plant Overview", "Bag Analysis", "Efficiency Metrics", "Forecasting"]
                )
                
                # Time period selection
                st.header("📅 Time Range")
                available_months = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
                start_month, end_month = st.select_slider(
                    'Select Date Range',
                    options=available_months,
                    value=(available_months[0], available_months[-1])
                )

            # Main content
            if analysis_type == "Plant Overview":
                # Calculate plant statistics
                plant_stats = calculate_plant_statistics(df, selected_plant)
                
                # Display metrics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("Total Usage", f"{plant_stats['Total Plant Usage']:,.0f}")
                with col2:
                    st.metric("Capacity Utilization", f"{plant_stats['Capacity Utilization']:.1f}%")
                with col3:
                    st.metric("Bag Types", str(plant_stats['Number of Bag Types']))
                with col4:
                    st.metric("MoM Growth", f"{plant_stats['MoM Growth Rate']:.1f}%")
                
                # Bag mix analysis
                st.subheader("📊 Bag Mix Analysis")
                fig_mix = px.pie(
                    values=list(plant_stats['Bag Mix'].values()),
                    names=list(plant_stats['Bag Mix'].keys()),
                    title="Product Mix Distribution"
                )
                st.plotly_chart(fig_mix, use_container_width=True)
                
                # Performance indicators
                st.subheader("📈 Performance Indicators")
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Usage Volatility", f"{plant_stats['Usage Volatility']:.1f}%")
                with col2:
                    if plant_stats['Seasonality Strength']:
                        st.metric("Seasonality", f"{plant_stats['Seasonality Strength']:.1f}%")
                    else:
                        st.metric("Seasonality", "Insufficient Data")

            elif analysis_type == "Bag Analysis":
                # Get bag patterns
                patterns = analyze_bag_patterns(df, selected_plant)
                
                # Stability analysis
                st.subheader("📊 Bag Stability Analysis")
                stability_df = pd.DataFrame.from_dict(
                    patterns['stability_scores'],
                    orient='index',
                    columns=['Stability Score']
                )
                st.bar_chart(stability_df)
                
                # Correlation matrix
                st.subheader("🔄 Bag Correlation Matrix")
                fig_corr = px.imshow(
                    patterns['correlations'],
                    title="Bag Demand Correlations",
                    color_continuous_scale="RdBu"
                )
                st.plotly_chart(fig_corr, use_container_width=True)
                
                # Complementary products
                if not patterns['complementary_products'].empty:
                    st.subheader("🤝 Complementary Products")
                    st.write(patterns['complementary_products'].head())

            elif analysis_type == "Efficiency Metrics":
                # Calculate efficiency metrics
                efficiency_metrics = calculate_efficiency_metrics(df, selected_plant)
                
                # Display metrics
                st.subheader("⚡ Efficiency Metrics")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Output Stability", f"{efficiency_metrics['output_stability']:.1f}%")
                with col2:
                    st.metric("Mix Complexity", f"{efficiency_metrics['mix_complexity']:.2f}")
                with col3:
                    st.metric("Capacity Consistency", f"{efficiency_metrics['capacity_consistency']:.1f}%")
                
                # Mix complexity trend
                st.subheader("📈 Mix Complexity Trend")
                mix_entropy_df = pd.DataFrame.from_dict(
                    efficiency_metrics['monthly_mix_entropy'],
                    orient='index',
                    columns=['Mix Entropy']
                )
                st.line_chart(mix_entropy_df)

            elif analysis_type == "Forecasting":
                # Bag selection for forecasting
                plant_bags = df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique()
                selected_bag = st.selectbox('Select Bag for Forecast:', sorted(plant_bags))
                
                # Create and display forecast plot
                forecast_fig = create_forecast_plot(df, selected_plant, selected_bag)
                st.plotly_chart(forecast_fig, use_container_width=True)

        except Exception as e:
            st.error(f"An error occurred: {str(e)}")
            st.write("Please check your data format and try again.")

if __name__ == '__main__':
    main()
