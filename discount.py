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

st.set_page_config(
    page_title="Discount Analytics Dashboard",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Updated CSS with optimized animation
st.markdown("""
<style>
    @keyframes ticker {
        0% { transform: translateX(100%); }
        100% { transform: translateX(-100%); }
    }
    
    .ticker-container {
        background-color: #0f172a;
        color: white;
        padding: 12px;
        overflow: hidden;
        white-space: nowrap;
        position: relative;
        margin-bottom: 20px;
        border-radius: 8px;
    }
    .ticker-content {
        display: inline-block;
        animation: ticker 2500s linear infinite;  /* Set to 60 seconds */
        animation-delay: -1250s;  /* Start halfway through to avoid initial wait */
        padding-right: 100%;
        will-change: transform;
        transform: translateZ(0);
    }
    
    .ticker-content:hover {
        animation-play-state: paused;
    }
    
    .ticker-item {
        display: inline-block;
        margin-right: 80px;
        font-size: 16px;
        padding: 5px 10px;
        opacity: 1;
        transition: opacity 0.3s;
    }
    
    .state-name {
        color: #10B981;
        font-weight: bold;
    }
    
    .month-name {
        color: #3B82F6;
        font-weight: bold;
    }
    
    .discount-value {
        color: #F59E0B;
    }

    @keyframes fadeIn {
        from { opacity: 0; }
        to { opacity: 1; }
    }

    .ticker-container {
        animation: fadeIn 0.5s ease-in;
    }

    .custom-card {
        background-color: white;
        padding: 1rem;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        margin-bottom: 1rem;
    }

    .custom-card h3 {
        color: #1e293b;
        margin-bottom: 0.5rem;
    }

    .custom-card p {
        margin: 0.5rem 0;
        color: #475569;
    }
</style>
""", unsafe_allow_html=True)

@st.cache_data(ttl=3600)
def process_excel_file(file_content, excluded_sheets):
    """Process Excel file and return processed data"""
    excel_data = io.BytesIO(file_content)
    excel_file = pd.ExcelFile(excel_data)
    processed_data = {}
    
    for sheet in excel_file.sheet_names:
        if not any(excluded_sheet in sheet for excluded_sheet in excluded_sheets):
            df = pd.read_excel(excel_data, sheet_name=sheet, usecols=range(22))
            
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
            
            processed_data[sheet] = df
            
    return processed_data

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
        """Create moving ticker with comprehensive discount information"""
        ticker_items = []
        
        # Get the last month (June in this case)
        last_month = "June"
        month_cols = self.month_columns[last_month]
        
        for state in data.keys():
            df = data[state]
            if not df.empty:
                state_text = f"<span class='state-name'>📍 {state}</span>"
                month_text = f"<span class='month-name'>📅 {last_month}</span>"
                
                discount_types = self.get_discount_types(df)
                discount_items = []
                
                for discount in discount_types:
                    mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == discount.strip()
                    filtered_df = df[mask]
                    
                    if len(filtered_df) > 0:
                        approved = filtered_df.iloc[0, month_cols['approved']]
                        actual = filtered_df.iloc[0, month_cols['actual']]
                        discount_items.append(
                            f"{discount}: <span class='discount-value'>₹{actual:,.2f}</span>"
                        )
                
                full_text = f"{state_text} | {month_text} | {' '.join(discount_items)}"
                ticker_items.append(f"<span class='ticker-item'>{full_text}</span>")
        
        # Repeat items 3 times for continuous flow
        ticker_items = ticker_items * 3
        
        ticker_html = f"""
        <div class="ticker-container">
            <div class="ticker-content">
                {' '.join(ticker_items)}
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
    def process_excel(self, uploaded_file):
        """Process uploaded Excel file using cached function"""
        return process_excel_file(uploaded_file.getvalue(), ['MP (U)', 'MP (JK)'])
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
        st.title("Dashboard Controls")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    # Main content
    st.title("Discount Analytics Dashboard")
    st.markdown("---")
    
    if uploaded_file is not None:
        # Use spinner to show loading state
        with st.spinner('Processing data...'):
            # Process data
            data = processor.process_excel(uploaded_file)
            
            # Create ticker immediately after data processing
            processor.create_ticker(data)
        
        # Rest of the dashboard components
        processor.create_summary_metrics(data)
        
        # State and discount selection
        col1, col2 = st.columns(2)
        with col1:
            selected_state = st.selectbox("Select State", list(data.keys()))
        
        if selected_state:
            with col2:
                discount_types = processor.get_discount_types(data[selected_state])
                selected_discount = st.selectbox("Select Discount Type", discount_types)
        
            processor.create_trend_chart(data, selected_state, selected_discount)
            
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
