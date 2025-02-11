import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
def main():
    st.title("Cement Plant Bag Usage Analysis",layout="wide")
    
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
            
            # Create the line graph using plotly
            fig = px.line(plot_df, x='Month', y=['Usage', 'Projected'],
                         title=f'Monthly Usage for {selected_bag} at {selected_plant}',
                         markers=True)
            
            # Customize the graph
            fig.update_traces(
                line_color='blue',
                name='Actual Usage',
                selector=dict(name='Usage')
            )
            fig.update_traces(
                line_color='red',
                line_dash='dash',
                name='Projected (Feb)',
                selector=dict(name='Projected')
            )
            
            # Update layout
            fig.update_layout(
                xaxis_title='Month',
                yaxis_title='Usage',
                legend_title='Type',
                hovermode='x'
            )
            
            # Display the graph
            st.plotly_chart(fig)
            
            # Display the raw data
            if st.checkbox('Show raw data'):
                st.write(plot_df)

if __name__ == '__main__':
    main()
