import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime

def format_date_for_display(date):
    """Convert datetime to 'MMM YYYY' format"""
    if isinstance(date, str):
        date = pd.to_datetime(date)
    return date.strftime('%b %Y')

def main():
    st.title("Cement Plant Bag Usage Analysis")
    
    # File uploader
    uploaded_file = st.file_uploader("Upload your Excel file", type=['xlsx', 'xls'])
    
    if uploaded_file is not None:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Remove the first column
        df = df.iloc[:, 1:]
        
        # Get unique plant names from "Cement Plant Sname" column
        unique_plants = sorted(df['Cement Plant Sname'].unique())
        
        # First dropdown for plant selection
        selected_plant = st.selectbox(
            'Select Cement Plant:',
            unique_plants
        )
        
        # Filter bags for selected plant
        plant_bags = df[df['Cement Plant Sname'] == selected_plant]['MAKTX'].unique()
        
        # Second dropdown for bag selection
        selected_bag = st.selectbox(
            'Select Bag:',
            sorted(plant_bags)
        )
        
        # Get the row for selected plant and bag
        selected_data = df[(df['Cement Plant Sname'] == selected_plant) & 
                         (df['MAKTX'] == selected_bag)]
        
        if not selected_data.empty:
            # Get all monthly columns (excluding 'Cement Plant Sname' and 'MAKTX')
            month_columns = [col for col in df.columns if col not in ['Cement Plant Sname', 'MAKTX']]
            
            # Create data for all months (for table)
            all_usage_data = []
            for month in month_columns:
                date = pd.to_datetime(month)
                usage = selected_data[month].iloc[0]
                all_usage_data.append({
                    'Date': date,
                    'Usage': usage
                })
            
            # Create DataFrame for all historical data
            all_data_df = pd.DataFrame(all_usage_data)
            all_data_df = all_data_df.sort_values('Date')
            
            # Add Month column for display
            all_data_df['Month'] = all_data_df['Date'].apply(format_date_for_display)
            
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
            
            # Create figure with custom layout
            fig = go.Figure()
            
            # Add actual usage line
            fig.add_trace(go.Scatter(
                x=plot_data['Month'],
                y=plot_data['Usage'],
                name='Actual Usage',
                line=dict(color='#1f77b4', width=3),
                mode='lines+markers',
                marker=dict(size=10, symbol='circle')
            ))
            
            # Add projected usage line for February 2025
            if 'Projected' in plot_data.columns:
                fig.add_trace(go.Scatter(
                    x=plot_data['Month'],
                    y=plot_data['Projected'],
                    name='Projected (Feb)',
                    line=dict(color='#ff7f0e', width=2, dash='dash'),
                    mode='lines'
                ))
            
            # Add brand rejuvenation vertical line
            fig.add_shape(
                type="line",
                x0="Jan 2025",
                x1="Jan 2025",
                y0=0,
                y1=plot_data['Usage'].max() * 1.1,
                line=dict(color="#FF4B4B", width=2, dash="dash"),
            )
            
            # Add annotation for brand rejuvenation
            fig.add_annotation(
                x="Jan 2025",
                y=plot_data['Usage'].max() * 1.15,
                text="Brand Rejuvenation<br>(15th Jan 2025)",
                showarrow=True,
                arrowhead=1,
                ax=0,
                ay=-40,
                font=dict(size=12, color="#FF4B4B"),
                bgcolor="white",
                bordercolor="#FF4B4B",
                borderwidth=2
            )
            
            # Add annotation for February data
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
                    bordercolor="#1f77b4",
                    borderwidth=2
                )
            
            # Update layout with enhanced styling
            fig.update_layout(
                title={
                    'text': f'Monthly Usage for {selected_bag} at {selected_plant}<br><sup>Showing data from April 2024 onwards</sup>',
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
                    gridcolor='lightgray',
                    tickangle=45
                ),
                yaxis=dict(
                    showgrid=True,
                    gridcolor='lightgray',
                    zeroline=True,
                    zerolinecolor='lightgray'
                ),
                legend=dict(
                    yanchor="top",
                    y=0.99,
                    xanchor="left",
                    x=0.01,
                    bgcolor='rgba(255, 255, 255, 0.8)'
                )
            )
            
            # Display the graph
            st.plotly_chart(fig, use_container_width=True)
            
            # Display the complete historical data
            st.subheader("Complete Historical Data")
            # Display the data without the Date column
            display_df = all_data_df[['Month', 'Usage']]
            st.dataframe(
                display_df.style.format({'Usage': '{:,.2f}'})
            )

if __name__ == '__main__':
    main()
