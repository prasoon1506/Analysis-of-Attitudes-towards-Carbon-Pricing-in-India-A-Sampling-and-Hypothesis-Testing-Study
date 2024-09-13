import streamlit as st
import pandas as pd
import io
import numpy as np
import seaborn as sns
import matplotlib.pyplot as plt

def process_dataframe(df):
    try:
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

        # Check if all expected columns are present
        missing_columns = [col for col in column_mapping.values() if col not in df.columns]
        if missing_columns:
            raise ValueError(f"Missing columns in the Excel file: {', '.join(missing_columns)}")

        df = df.rename(columns=column_mapping)

        # Ensure all data columns are numeric
        for col in df.columns:
            if col not in ['Region', 'Dist Name']:
                df[col] = pd.to_numeric(df[col], errors='coerce')

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
        for col in df.columns:
            if 'Growth/Degrowth' in col:
                df[col] = df[col].replace([np.inf, -np.inf], np.nan)

        return df

    except Exception as e:
        st.error(f"Error processing the dataframe: {str(e)}")
        st.error("Please check your Excel file for data consistency and try again.")
        return None

def validate_data_for_graphs(df):
    numeric_columns = df.select_dtypes(include=[np.number]).columns
    for col in numeric_columns:
        if df[col].isnull().any():
            st.warning(f"Column {col} contains null values")
        if np.isinf(df[col]).any():
            st.warning(f"Column {col} contains infinite values")
        if not np.issubdtype(df[col].dtype, np.number):
            st.warning(f"Column {col} contains non-numeric values")

def create_graph(df_23, df_24, channel_col, entity_name):
    try:
        months = ['Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep']
        
        fig, ax = plt.subplots(figsize=(10, 6))
        
        # Extract values for 2023 and 2024
        values_23 = df_23.iloc[0, 1:].values
        values_24 = df_24.iloc[0, 1:].values
        
        # Check for NaN or infinite values
        if np.isnan(values_23).any() or np.isinf(values_23).any() or np.isnan(values_24).any() or np.isinf(values_24).any():
            raise ValueError("Data contains NaN or infinite values")
        
        # Check if all values are numeric
        if not (np.issubdtype(values_23.dtype, np.number) and np.issubdtype(values_24.dtype, np.number)):
            raise ValueError("Data contains non-numeric values")
        
        # Plot 2023 data
        sns.lineplot(x=months, y=values_23, marker='o', label=f'FY23{channel_col}', ax=ax)
        
        # Plot 2024 data (stop at August)
        sns.lineplot(x=months[:-1], y=values_24, marker='o', label=f'FY24{channel_col}', ax=ax)

        ax.set_title(f'Sales Data by Month for {entity_name} {channel_col}')
        ax.set_xlabel('Month')
        ax.set_ylabel('Quantity Sold')
        ax.legend()
        
        return fig
    except Exception as e:
        st.error(f"Error creating graph for {entity_name} {channel_col}: {str(e)}")
        st.error(f"Data for 2023: {df_23.iloc[0, 1:].values}")
        st.error(f"Data for 2024: {df_24.iloc[0, 1:].values}")
        return None

def main():
    st.title('Sales Data Analysis')

    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
    if uploaded_file is not None:
        try:
            df = pd.read_excel(uploaded_file)
            processed_df = process_dataframe(df)
            
            if processed_df is None:
                return

            st.write(f"Uploaded file: {uploaded_file.name}")

            aggregation = st.selectbox('Aggregate by:', ['District', 'Region'])

            if aggregation == 'Region':
                selected_regions = st.multiselect('Region:', processed_df['Region'].unique())
                filtered_data = processed_df[processed_df['Region'].isin(selected_regions)].copy()
                grouped_data = filtered_data.groupby('Region').sum().reset_index()
                entity_col = 'Region'
            else:
                selected_districts = st.multiselect('District:', processed_df['Dist Name'].unique())
                filtered_data = processed_df[processed_df['Dist Name'].isin(selected_districts)].copy()
                grouped_data = filtered_data.copy()
                entity_col = 'Dist Name'

            selected_channels = st.multiselect('Channel:', ['Overall', 'Trade', 'Non-Trade'])

            if st.button('Run Analysis'):
                if grouped_data.empty:
                    st.warning("No data to display. Please select at least one region or district.")
                    return

                validate_data_for_graphs(grouped_data)

                # Calculate growth rates and quarterly requirements
                for prefix in ['', ' Trade', ' Non-Trade']:
                    grouped_data[f'Growth/Degrowth(MTD){prefix}'] = (grouped_data[f'24-Aug{prefix}'] - grouped_data[f'23-Aug{prefix}']) / grouped_data[f'23-Aug{prefix}'] * 100
                    grouped_data[f'Growth/Degrowth(YTD){prefix}'] = (grouped_data[f'FY 2024 till Aug{prefix}'] - grouped_data[f'FY 2023 till Aug{prefix}']) / grouped_data[f'FY 2023 till Aug{prefix}'] * 100
                    grouped_data[f'Quarterly Requirement{prefix}'] = grouped_data[f'Q3 2023{prefix}'] - grouped_data[f'Q3 2024 till August{prefix}']

                for selected_channel in selected_channels:
                    if selected_channel == 'Trade':
                        suffix = ' Trade'
                    elif selected_channel == 'Non-Trade':
                        suffix = ' Non-Trade'
                    else:
                        suffix = ''
                    
                    columns_to_display = [entity_col, f'24-Aug{suffix}', f'23-Aug{suffix}', f'Growth/Degrowth(MTD){suffix}',
                                          f'FY 2024 till Aug{suffix}', f'FY 2023 till Aug{suffix}', f'Growth/Degrowth(YTD){suffix}',
                                          f'Q3 2023{suffix}', f'Q3 2024 till August{suffix}', f'Quarterly Requirement{suffix}']
                    
                    temp_df = grouped_data[columns_to_display].copy()
                    
                    for col in temp_df.columns:
                        if 'Growth/Degrowth' in col:
                            temp_df[col] = temp_df[col].apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else '')
                        elif col != entity_col:
                            temp_df[col] = pd.to_numeric(temp_df[col], errors='coerce').round().astype('Int64')
                    
                    st.subheader(f"{selected_channel} Data")
                    st.dataframe(temp_df.style.applymap(lambda x: 'color: green' if ('%' in str(x) and float(str(x).replace('%', '')) > 0) else 'color: red', subset=[col for col in temp_df.columns if 'Growth/Degrowth' in col]))

                # Display graphs
                if selected_channels:
                    months_23 = ['23-Apr', '23-May', '23-Jun', '23-Jul', '23-Aug', '23-Sep']
                    months_24 = ['24-Apr', '24-May', '24-Jun', '24-Jul', '24-Aug']
                    
                    entities = grouped_data[entity_col].unique()
                    
                    for entity in entities:
                        entity_data = grouped_data[grouped_data[entity_col] == entity]
                        
                        df_23 = entity_data[months_23]
                        df_24 = entity_data[months_24]

                        for channel in selected_channels:
                            try:
                                if channel == 'Overall':
                                    fig = create_graph(df_23, df_24, '', entity)
                                elif channel == 'Trade':
                                    df_23_trade = entity_data[[col + ' Trade' for col in months_23]]
                                    df_24_trade = entity_data[[col + ' Trade' for col in months_24]]
                                    fig = create_graph(df_23_trade, df_24_trade, ' Trade', entity)
                                elif channel == 'Non-Trade':
                                    df_23_nontrade = entity_data[[col + ' Non-Trade' for col in months_23]]
                                    df_24_nontrade = entity_data[[col + ' Non-Trade' for col in months_24]]
                                    fig = create_graph(df_23_nontrade, df_24_nontrade, ' Non-Trade', entity)
                                
                                if fig is not None:
                                    st.pyplot(fig)
                                    plt.close(fig)
                            except Exception as e:
                               st.error(f"Error processing graph for {entity} - {channel}: {str(e)}")
if __name__ == "__main__":
    main()
