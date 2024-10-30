import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
from datetime import datetime
import time
import streamlit.components.v1 as components
import io
import warnings
warnings.filterwarnings('ignore')

# Set page config
st.set_page_config(
    page_title="Discount Analytics Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for styling
st.markdown("""
<style>
    /* Main container */
    .main {
        background-color: #f0f2f6;
    }
    
    /* Headers */
    .css-10trblm {
        color: #1f2937;
        font-weight: 600;
    }
    
    /* Metric cards */
    .css-1r6slb0 {
        background-color: white;
        border-radius: 10px;
        padding: 1rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: #1f2937;
    }
    
    /* File uploader */
    .stFileUploader {
        padding: 1rem;
        border-radius: 10px;
        background-color: white;
    }
    
    /* Ticker animation */
    @keyframes ticker {
        0% { transform: translateX(100%); }
        100% { transform: translateX(-100%); }
    }
    
    .ticker-container {
        background-color: #0f172a;
        color: white;
        padding: 10px;
        overflow: hidden;
        white-space: nowrap;
    }
    
    .ticker-content {
        display: inline-block;
        animation: ticker 20s linear infinite;
    }
    
    /* Custom card */
    .custom-card {
        background-color: white;
        border-radius: 10px;
        padding: 20px;
        margin: 10px 0;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
</style>
""", unsafe_allow_html=True)

class DiscountAnalytics:
    def __init__(self):
        self.excluded_discounts = [
            'Sub Total',
            'TOTAL OF DP PAYOUT',
            'TOTAL OF STS & RD',
            'Other (Please specify',
            'G. TOTAL'
        ]
        
        self.discount_mappings = {
            'group1': {
                'states': ['HP', 'JMU', 'PUN'],
                'discounts': ['CASH DISCOUNT', 'ADVANCE CD & NIL OS']
            },
            'group2': {
                'states': ['UP (W)'],
                'discounts': ['CD', 'Adv CD']
            }
        }
        
        self.combined_discount_name = 'CD and Advance CD'
        
        self.month_columns = {
            'April': {
                'quantity': 1,
                'approved': 2,
                'actual': 4
            },
            'May': {
                'quantity': 8,
                'approved': 9,
                'actual': 11
            },
            'June': {
                'quantity': 15,
                'approved': 16,
                'actual': 18
            }
        }
        
    def create_ticker(self, data):
        """Create moving ticker with discount rates"""
        ticker_items = []
        for state in data.keys():
            df = data[state]
            if not df.empty:
                ticker_items.append(f"{state}: ₹{df.iloc[0, 4]:,.2f}")
        
        ticker_html = f"""
        <div class="ticker-container">
            <div class="ticker-content">
                {'  |  '.join(ticker_items * 2)}
            </div>
        </div>
        """
        st.markdown(ticker_html, unsafe_allow_html=True)

    def create_summary_metrics(self, data):
        """Create summary metrics cards"""
        total_states = len(data)
        total_discounts = sum(len(self.get_discount_types(df)) for df in data.values())
        avg_discount = np.mean([df.iloc[0, 4] for df in data.values() if not df.empty])
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric("Total States", total_states, "Active")
        with col2:
            st.metric("Total Discount Types", total_discounts, "Available")
        with col3:
            st.metric("Average Discount Rate", f"₹{avg_discount:,.2f}", "Per Bag")

    def create_trend_chart(self, data, selected_state, selected_discount):
        """Create trend chart using Plotly"""
        df = data[selected_state]
        
        if selected_discount == self.combined_discount_name:
            monthly_data = {
                month: self.get_combined_data(df, cols, selected_state)
                for month, cols in self.month_columns.items()
            }
        else:
            mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == selected_discount.strip()
            filtered_df = df[mask]
            if len(filtered_df) > 0:
                monthly_data = {
                    month: {
                        'actual': filtered_df.iloc[0, cols['actual']],
                        'approved': filtered_df.iloc[0, cols['approved']]
                    }
                    for month, cols in self.month_columns.items()
                }
        
        months = list(monthly_data.keys())
        actual_values = [data['actual'] for data in monthly_data.values()]
        approved_values = [data['approved'] for data in monthly_data.values()]
        
        fig = go.Figure()
        
        fig.add_trace(go.Scatter(
            x=months,
            y=actual_values,
            name='Actual',
            line=dict(color='#10B981', width=3)
        ))
        
        fig.add_trace(go.Scatter(
            x=months,
            y=approved_values,
            name='Approved',
            line=dict(color='#3B82F6', width=3)
        ))
        
        fig.update_layout(
            title=f'Discount Trends - {selected_state}',
            xaxis_title='Month',
            yaxis_title='Discount Rate (₹/Bag)',
            template='plotly_white',
            height=400,
            margin=dict(t=50, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig, use_container_width=True)

    @st.cache_data
    def process_excel(self, uploaded_file):
        """Process uploaded Excel file"""
        excel_data = io.BytesIO(uploaded_file.getvalue())
        excel_file = pd.ExcelFile(excel_data)
        processed_data = {}
        
        for sheet in excel_file.sheet_names:
            if not self.should_process_sheet(sheet):
                continue
                
            df = pd.read_excel(excel_data, sheet_name=sheet, usecols=range(22))
            df = self.preprocess_dataframe(df)
            processed_data[sheet] = df
            
        return processed_data

    def should_process_sheet(self, sheet_name):
        """Check if sheet should be processed"""
        excluded_sheets = ['MP (U)', 'MP (JK)']
        return not any(excluded_sheet in sheet_name for excluded_sheet in excluded_sheets)

    def preprocess_dataframe(self, df):
        """Preprocess the dataframe"""
        # Find start index
        cash_discount_patterns = ['CASH DISCOUNT', 'Cash Discount', 'CD']
        start_idx = None
        
        for idx, value in enumerate(df.iloc[:, 0]):
            if isinstance(value, str):
                if any(pattern.lower() in value.lower() for pattern in cash_discount_patterns):
                    start_idx = idx
                    break
        
        if start_idx is not None:
            df = df.iloc[start_idx:].reset_index(drop=True)
        
        # Trim at G. Total
        g_total_idx = None
        for idx, value in enumerate(df.iloc[:, 0]):
            if isinstance(value, str) and 'G. TOTAL' in value:
                g_total_idx = idx
                break
        
        if g_total_idx is not None:
            df = df.iloc[:g_total_idx].copy()
        
        return df

    def get_discount_types(self, df):
        """Get unique discount types"""
        first_col = df.iloc[:, 0]
        return sorted([
            d for d in first_col.unique()
            if isinstance(d, str) and d.strip() not in self.excluded_discounts
        ])

    def get_combined_data(self, df, month_cols, state):
        """Get combined discount data"""
        combined_data = {'actual': np.nan, 'approved': np.nan}
        
        state_group = next(
            (group for group, config in self.discount_mappings.items()
             if state in config['states']),
            None
        )
        
        if state_group:
            relevant_discounts = self.discount_mappings[state_group]['discounts']
            mask = df.iloc[:, 0].fillna('').astype(str).str.strip().isin(relevant_discounts)
            filtered_df = df[mask]
            
            if len(filtered_df) > 0:
                combined_data['approved'] = filtered_df.iloc[:, month_cols['approved']].sum()
                combined_data['actual'] = filtered_df.iloc[:, month_cols['actual']].sum()
        
        return combined_data

def main():
    # Initialize the processor
    processor = DiscountAnalytics()
    
    # Sidebar
    with st.sidebar:
        st.image("https://via.placeholder.com/150x50.png?text=Logo", use_column_width=True)
        st.title("Dashboard Controls")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    # Main content
    st.title("Discount Analytics Dashboard")
    st.markdown("---")
    
    if uploaded_file:
        # Process data
        data = processor.process_excel(uploaded_file)
        
        # Create ticker
        processor.create_ticker(data)
        
        # Summary metrics
        processor.create_summary_metrics(data)
        
        # State and discount selection
        col1, col2 = st.columns(2)
        with col1:
            selected_state = st.selectbox("Select State", list(data.keys()))
        
        if selected_state:
            with col2:
                discount_types = processor.get_discount_types(data[selected_state])
                selected_discount = st.selectbox("Select Discount Type", discount_types)
        
            # Create trend chart
            processor.create_trend_chart(data, selected_state, selected_discount)
            
            # Display monthly details
            st.subheader("Monthly Details")
            cols = st.columns(3)
            
            for idx, (month, month_cols) in enumerate(processor.month_columns.items()):
                with cols[idx]:
                    st.markdown(f"""
                    <div class="custom-card">
                        <h3>{month}</h3>
                        <p><strong>Quantity Sold:</strong> {data[selected_state].iloc[0, month_cols['quantity']]:,.2f}</p>
                        <p><strong>Approved Rate:</strong> ₹{data[selected_state].iloc[0, month_cols['approved']]:,.2f}</p>
                        <p><strong>Actual Rate:</strong> ₹{data[selected_state].iloc[0, month_cols['actual']]:,.2f}</p>
                    </div>
                    """, unsafe_allow_html=True)
    
    else:
        st.info("Please upload an Excel file to begin analysis.")
        
        # Placeholder metrics for demo
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total States", "0", "Waiting")
        with col2:
            st.metric("Total Discount Types", "0", "Waiting")
        with col3:
            st.metric("Average Discount Rate", "₹0.00", "Waiting")

if __name__ == "__main__":
    main()
