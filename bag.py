import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
import plotly.express as px

def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def calculate_statistics(data_df):
    """Calculate key statistics from the usage data"""
    stats = {
        'Total Usage': data_df['Usage'].sum(),
        'Average Monthly Usage': data_df['Usage'].mean(),
        'Highest Usage': data_df['Usage'].max(),
        'Lowest Usage': data_df['Usage'].min(),
        'Usage Variance': data_df['Usage'].var(),
        'Month-over-Month Change': (data_df['Usage'].iloc[-1] - data_df['Usage'].iloc[-2]) / data_df['Usage'].iloc[-2] * 100
    }
    return stats

def create_year_over_year_comparison(data_df):
    """Create year-over-year comparison data"""
    data_df['Year'] = data_df['Date'].dt.year
    data_df['Month'] = data_df['Date'].dt.month
    yearly_comparison = data_df.pivot(index='Month', columns='Year', values='Usage')
    return yearly_comparison

def main():
    # Set page configuration with custom theme
    st.set_page_config(
        page_title="Cement Plant Bag Usage Analysis",
        layout='wide',
        initial_sidebar_state='expanded'
    )

    # Custom CSS for better styling
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
        </style>
    """, unsafe_allow_html=True)

    # Title with custom styling
    st.title("📊 Cement Plant Bag Usage Analysis")
    
    # File uploader in sidebar
    with st.sidebar:
        st.header("📁 Data Input")
        uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Read and process the Excel file
        df = pd.read_excel(uploaded_file)
        df = df.iloc[:, 1:]  # Remove the first column
        
        # Sidebar filters
        with st.sidebar:
            st.header("🎯 Filters")
            unique_plants = sorted(df['Cement Plant Sname'].unique())
            selected_plant = st.selectbox('Select Cement Plant:', unique_plants)
            
            plant_bags = df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique()
            selected_bag = st.selectbox('Select Bag:', sorted(plant_bags))

        # Get selected data
        selected_data = df[(df['Cement Plant Sname'] == selected_plant) & 
                         (df['MAKTX'] == selected_bag)]
        
        if not selected_data.empty:
            # Process data
            month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
            all_usage_data = []
            for month in month_columns:
                date = pd.to_datetime(month)
                usage = selected_data[month].iloc[0]
                all_usage_data.append({
                    'Date': date,
                    'Usage': usage
                })
            
            # Create and process DataFrames
            all_data_df = pd.DataFrame(all_usage_data)
            all_data_df = all_data_df.sort_values('Date')
            all_data_df['Month'] = all_data_df['Date'].apply(format_date_for_display)
            
            # Calculate statistics
            stats = calculate_statistics(all_data_df)
            
            # Display key metrics in columns
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Total Usage", f"{stats['Total Usage']:,.0f}")
            with col2:
                st.metric("Average Monthly", f"{stats['Average Monthly Usage']:,.0f}")
            with col3:
                st.metric("Highest Usage", f"{stats['Highest Usage']:,.0f}")
            with col4:
                st.metric("MoM Change", f"{stats['Month-over-Month Change']:,.1f}%")

            # Create tabs for different visualizations
            tab1, tab2, tab3 = st.tabs(["📈 Usage Trend", "📊 Year Comparison", "📑 Historical Data"])
            
            with tab1:
                # Filter data from Apr 2024 onwards for plotting
                apr_2024_date = pd.to_datetime('2024-04-01')
                plot_data = all_data_df[all_data_df['Date'] >= apr_2024_date].copy()
                
                # Add projected data for February 2025
                if any(plot_data['Date'].dt.strftime('%Y-%m') == '2025-02'):
                    feb_data = plot_data[plot_data['Date'].dt.strftime('%Y-%m') == '2025-02']
                    feb_usage = feb_data['Usage'].iloc[0]
                    daily_avg = feb_usage / 9
                    projected_usage = daily_avg * 29
                    plot_data.loc[plot_data['Date'].dt.strftime('%Y-%m') == '2025-02', 'Projected'] = projected_usage
                
                # Create enhanced figure
                fig = go.Figure()
                
                # Add actual usage line with improved styling
                fig.add_trace(go.Scatter(
                    x=plot_data['Month'],
                    y=plot_data['Usage'],
                    name='Actual Usage',
                    line=dict(color='#2E86C1', width=3),
                    mode='lines+markers',
                    marker=dict(size=10, symbol='circle')
                ))
                
                # Add projected usage line
                if 'Projected' in plot_data.columns:
                    fig.add_trace(go.Scatter(
                        x=plot_data['Month'],
                        y=plot_data['Projected'],
                        name='Projected (Feb)',
                        line=dict(color='#E67E22', width=2, dash='dash'),
                        mode='lines'
                    ))
                
                # Add brand rejuvenation marker
                fig.add_shape(
                    type="line",
                    x0="Jan 2025",
                    x1="Jan 2025",
                    y0=0,
                    y1=plot_data['Usage'].max() * 1.1,
                    line=dict(color="#E74C3C", width=2, dash="dash"),
                )
                
                # Enhanced annotations
                fig.add_annotation(
                    x="Jan 2025",
                    y=plot_data['Usage'].max() * 1.15,
                    text="Brand Rejuvenation<br>(15th Jan 2025)",
                    showarrow=True,
                    arrowhead=1,
                    ax=0,
                    ay=-40,
                    font=dict(size=12, color="#E74C3C"),
                    bgcolor="white",
                    bordercolor="#E74C3C",
                    borderwidth=2
                )
                
                if any(plot_data['Month'] == 'Feb 2025'):
                    feb_data = plot_data[plot_data['Month'] == 'Feb 2025']
                    fig.add_annotation(
                        x="Feb 2025",
                        y=feb_data['Usage'].iloc[0],
                        text="Till 9th Feb",
                        showarrow=True,
                        arrowhead=1,
                        ax=0,
                        ay=-40,
                        font=dict(size=12),
                        bgcolor="white",
                        bordercolor="#2E86C1",
                        borderwidth=2
                    )
                
                # Enhanced layout
                fig.update_layout(
                    title={
                        'text': f'Monthly Usage Trend for {selected_bag}<br><sup>{selected_plant}</sup>',
                        'y':0.95,
                        'x':0.5,
                        'xanchor': 'center',
                        'yanchor': 'top',
                        'font': dict(size=20)
                    },
                    xaxis_title='Month',
                    yaxis_title='Usage',
                    legend_title='Type',
                    hovermode='x unified',
                    plot_bgcolor='white',
                    paper_bgcolor='white',
                    showlegend=True,
                    xaxis=dict(
                        showgrid=True,
                        gridcolor='rgba(0,0,0,0.1)',
                        tickangle=45
                    ),
                    yaxis=dict(
                        showgrid=True,
                        gridcolor='rgba(0,0,0,0.1)',
                        zeroline=True,
                        zerolinecolor='rgba(0,0,0,0.2)'
                    ),
                    legend=dict(
                        yanchor="top",
                        y=0.99,
                        xanchor="left",
                        x=0.01,
                        bgcolor='rgba(255, 255, 255, 0.8)'
                    )
                )
                
                st.plotly_chart(fig, use_container_width=True)

            with tab2:
                # Year over year comparison
                yearly_comparison = create_year_over_year_comparison(all_data_df)
                
                # Create heatmap for year comparison
                fig_heatmap = px.imshow(
                    yearly_comparison,
                    labels=dict(x="Year", y="Month", color="Usage"),
                    aspect="auto",
                    color_continuous_scale="RdYlBu_r"
                )
                
                fig_heatmap.update_layout(
                    title="Year-over-Year Usage Comparison",
                    xaxis_title="Year",
                    yaxis_title="Month",
                )
                
                st.plotly_chart(fig_heatmap, use_container_width=True)

            with tab3:
                # Enhanced historical data display
                st.subheader("📜 Complete Historical Data")
                
                # Add percentage change column
                display_df = all_data_df[['Month', 'Usage']].copy()
                display_df['% Change'] = display_df['Usage'].pct_change() * 100
                
                # Style the dataframe
                st.dataframe(
                    display_df.style
                    .format({
                        'Usage': '{:,.2f}',
                        '% Change': '{:+.2f}%'
                    })
                    .background_gradient(subset=['Usage'], cmap='Blues')
                    .background_gradient(subset=['% Change'], cmap='RdYlGn')
                )

if __name__ == '__main__':
    main()
