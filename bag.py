import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime

def main():
    st.set_page_config(page_title="Cement Plant Bag Usage", layout="wide")
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
            # Extract monthly data
            months = ['Apr', 'May', 'Jun', 'July', 'Aug', 'Sep', 
                     'Oct', 'Nov', 'Dec', 'Jan', 'Feb']
            
            # Create data for plotting
            usage_data = []
            for month in months:
                usage = selected_data[month].iloc[0]
                # For February 2025, prorate the data
                if month == 'Feb':
                    # Calculate daily average and project for full month
                    daily_avg = usage / 9  # Usage till 9th Feb
                    projected_usage = daily_avg * 29  # February 2025 has 29 days
                    usage_data.append({
                        'Month': f"{month} 2025" if month in ['Jan', 'Feb'] else f"{month} 2024",
                        'Usage': usage,
                        'Projected': projected_usage if month == 'Feb' else None
                    })
                else:
                    usage_data.append({
                        'Month': f"{month} 2025" if month in ['Jan', 'Feb'] else f"{month} 2024",
                        'Usage': usage,
                        'Projected': None
                    })
            
            # Create DataFrame for plotting
            plot_df = pd.DataFrame(usage_data)
            
            # Create figure with custom layout
            fig = go.Figure()
            
            # Add actual usage line
            fig.add_trace(go.Scatter(
                x=plot_df['Month'],
                y=plot_df['Usage'],
                name='Actual Usage',
                line=dict(color='#1f77b4', width=3),
                mode='lines+markers',
                marker=dict(size=10, symbol='circle')
            ))
            
            # Add projected usage line for February
            fig.add_trace(go.Scatter(
                x=plot_df['Month'],
                y=plot_df['Projected'],
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
                y1=plot_df['Usage'].max() * 1.1,
                line=dict(color="#FF4B4B", width=2, dash="dash"),
            )
            
            # Add annotation for brand rejuvenation
            fig.add_annotation(
                x="Jan 2025",
                y=plot_df['Usage'].max() * 1.15,
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
            fig.add_annotation(
                x="Feb 2025",
                y=plot_df.loc[plot_df['Month'] == 'Feb 2025', 'Usage'].iloc[0],
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
                    'text': f'Monthly Usage for {selected_bag} at {selected_plant}',
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
            
            # Display the raw data
            if st.checkbox('Show raw data'):
                st.write(plot_df)

if __name__ == '__main__':
    main()
