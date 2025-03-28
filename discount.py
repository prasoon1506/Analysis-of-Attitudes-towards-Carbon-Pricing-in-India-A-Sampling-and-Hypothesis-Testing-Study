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
st.set_page_config(page_title="Discount Analytics Dashboard",page_icon="💰",layout="wide",initial_sidebar_state="expanded")
st.markdown("""<style>/* Global Styles */[data-testid="stSidebar"] {background-color: #f8fafc;border-right: 1px solid #e2e8f0;}.stButton button {background-color: #3b82f6;color: white;border-radius: 6px;padding: 0.5rem 1rem;border: none;transition: all 0.2s;}.stButton button:hover {background-color: #2563eb;box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);}/* Ticker Animation */@keyframes ticker {0% { transform: translateX(100%); }100% { transform: translateX(-100%); }}.ticker-container {background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);color: white;padding: 16px;overflow: hidden;
        white-space: nowrap;
        position: relative;
        margin-bottom: 24px;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
    }
    
    .ticker-content {
        display: inline-block;
        animation: ticker 2500s linear infinite;
        animation-delay: -1250s;
        padding-right: 100%;
        will-change: transform;
    }
    
    .ticker-content:hover {
        animation-play-state: paused;
    }
    
    .ticker-item {
        display: inline-block;
        margin-right: 80px;
        font-size: 16px;
        padding: 8px 16px;
        opacity: 1;
        transition: opacity 0.3s;
        background: rgba(255, 255, 255, 0.1);
        border-radius: 8px;
    }
    
    /* Enhanced Metrics */
    .state-name {
        color: #10B981;
        font-weight: 600;
    }
    
    .month-name {
        color: #60A5FA;
        font-weight: 600;
    }
    
    .discount-value {
        color: #FBBF24;
        font-weight: 600;
    }
    
    /* Card Styles */
    .metric-card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        transition: transform 0.2s;
        border: 1px solid #e2e8f0;
    }
    
    .metric-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 10px 15px -3px rgba(0, 0, 0, 0.1);
    }
    
    .metric-value {
        font-size: 2rem;
        font-weight: 600;
        color: #1e293b;
    }
    
    .metric-label {
        color: #64748b;
        font-size: 0.875rem;
        margin-top: 0.5rem;
    }
    
    /* Chart Container */
    .chart-container {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);
        margin: 1rem 0;
        border: 1px solid #e2e8f0;
    }
    
    /* Selectbox Styling */
    .stSelectbox {
        background: white;
        border-radius: 8px;
        border: 1px solid #e2e8f0;
    }
    
    /* Custom Header */
    .dashboard-header {
        padding: 1.5rem;
        background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
        color: white;
        border-radius: 12px;
        margin-bottom: 2rem;
        text-align: center;
    }
    
    .dashboard-title {
        font-size: 2rem;
        font-weight: 600;
        margin-bottom: 0.5rem;
    }
    
    .dashboard-subtitle {
        color: #94a3b8;
        font-size: 1rem;
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
        self.total_patterns = ['G. TOTAL', 'G.TOTAL', 'G. Total', 'G.Total', 'GRAND TOTAL',"G. Total (STD + STS)"]
        self.excluded_states = ['MP (JK)', 'MP (U)','East']
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
            
            # Get state group for combined discounts
            state_group = next(
                (group for group, config in self.discount_mappings.items()
                 if state in config['states']),
                None
            )
            
            discount_items = []
            
            if state_group:
                # Handle combined discounts
                relevant_discounts = self.discount_mappings[state_group]['discounts']
                combined_data = self.get_combined_data(df, month_cols, state)
                
                if combined_data:
                    actual = combined_data.get('actual', 0)
                    discount_items.append(
                        f"{self.combined_discount_name}: <span class='discount-value'>₹{actual:,.2f}</span>"
                    )
                
                # Add other non-combined discounts
                for discount in self.get_discount_types(df, state):
                    if discount != self.combined_discount_name:
                        mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == discount.strip()
                        filtered_df = df[mask]
                        if len(filtered_df) > 0:
                            actual = filtered_df.iloc[0, month_cols['actual']]
                            discount_items.append(
                                f"{discount}: <span class='discount-value'>₹{actual:,.2f}</span>"
                            )
            else:
                # Normal processing for states without combined discounts
                for discount in self.get_discount_types(df, state):
                    mask = df.iloc[:, 0].fillna('').astype(str).str.strip() == discount.strip()
                    filtered_df = df[mask]
                    if len(filtered_df) > 0:
                        actual = filtered_df.iloc[0, month_cols['actual']]
                        discount_items.append(
                            f"{discount}: <span class='discount-value'>₹{actual:,.2f}</span>"
                        )
            
            if discount_items:
                full_text = f"{state_text} | {month_text} | {' | '.join(discount_items)}"
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
    def create_monthly_metrics(self, data, selected_state, selected_discount):
        """Create monthly metrics based on selected discount type"""
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
                        'approved': filtered_df.iloc[0, cols['approved']],
                        'quantity': filtered_df.iloc[0, cols['quantity']]
                    }
                    for month, cols in self.month_columns.items()
                }
        
        # Create three columns for each month
        for month, data in monthly_data.items():
            st.markdown(f"""
            <div style='text-align: center; margin-bottom: 10px;'>
                <h3 style='color: #1e293b; margin-bottom: 15px;'>{month}</h3>
            </div>
            """, unsafe_allow_html=True)
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                quantity = data.get('quantity', 0)
                st.metric(
                    "Quantity Sold",
                    f"{quantity:,.2f}",
                    delta=None,
                    help=f"Total quantity sold in {month}"
                )
            
            with col2:
                approved = data.get('approved', 0)
                st.metric(
                    "Approved Payout",
                    f"₹{approved:,.2f}",
                    delta=None,
                    help=f"Approved discount rate for {month}"
                )
            
            with col3:
                actual = data.get('actual', 0)
                difference = approved - actual
                delta_color = "normal" if difference >= 0 else "inverse"
                st.metric(
                    "Actual Payout",
                    f"₹{actual:,.2f}",
                    delta=f"₹{abs(difference):,.2f}" + (" under approved" if difference >= 0 else " over approved"),
                    delta_color=delta_color,
                    help=f"Actual discount rate for {month}"
                )
            
            st.markdown("---")
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
        
        # Create and display the difference chart
        self.create_difference_chart(months, approved_values, actual_values, selected_state)

    def create_difference_chart(self, months, approved_values, actual_values, selected_state):
        """Create chart showing difference between approved and actual rates"""
        differences = [approved - actual for approved, actual in zip(approved_values, actual_values)]
        
        fig = go.Figure()
        
        # Add separate traces for positive and negative differences
        for i in range(len(months)):
            color = '#10B981' if differences[i] >= 0 else '#EF4444'  # Green for positive, red for negative
            fig.add_trace(go.Scatter(
                x=[months[i], months[i]],
                y=[0, differences[i]],
                mode='lines',
                line=dict(color=color, width=3),
                showlegend=False
            ))
        
        # Add markers at the difference points
        fig.add_trace(go.Scatter(
            x=months,
            y=differences,
            mode='markers',
            marker=dict(
                size=8,
                color=['#10B981' if d >= 0 else '#EF4444' for d in differences],
                line=dict(width=2, color='white')
            ),
            name='Difference'
        ))
        
        # Add a horizontal line at y=0
        fig.add_shape(
            type='line',
            x0=months[0],
            x1=months[-1],
            y0=0,
            y1=0,
            line=dict(color='gray', width=1, dash='dash')
        )
        
        fig.update_layout(
            title=f'Approved vs Actual Difference - {selected_state}',
            xaxis_title='Month',
            yaxis_title='Difference in Discount Rate (₹/Bag)',
            template='plotly_white',
            height=300,
            margin=dict(t=50, b=50, l=50, r=50)
        )
        
        st.plotly_chart(fig, use_container_width=True)
    def get_discount_types(self, df, state=None):
     first_col = df.iloc[:, 0]
     valid_discounts = []
     if state:
        state_group = next(
            (group for group, config in self.discount_mappings.items()
             if state in config['states']),
            None
        )
        
        if state_group:
            # Get the relevant discounts for this state
            relevant_discounts = self.discount_mappings[state_group]['discounts']
            
            # Add the combined discount name if any of the discounts to combine exist
            if any(d in first_col.values for d in relevant_discounts):
                valid_discounts.append(self.combined_discount_name)
            
            # Add other discounts that aren't being combined
            for d in first_col.unique():
                if (isinstance(d, str) and 
                    d.strip() not in self.excluded_discounts and 
                    d.strip() not in relevant_discounts):
                    valid_discounts.append(d)
        else:
            # Normal processing for other states
            valid_discounts = [
                d for d in first_col.unique() 
                if isinstance(d, str) and d.strip() not in self.excluded_discounts
            ]
     else:
        # When no state is provided (for ticker), return all unique discounts
        valid_discounts = [
            d for d in first_col.unique() 
            if isinstance(d, str) and d.strip() not in self.excluded_discounts
        ]
    
     return sorted(valid_discounts)
    def get_combined_data(self, df, month_cols, state):
     combined_data = {
        'actual': np.nan, 
        'approved': np.nan,
        'quantity': np.nan
    }
    
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
            # Sum up the values for all relevant discounts
            combined_data['approved'] = filtered_df.iloc[:, month_cols['approved']].sum()
            combined_data['actual'] = filtered_df.iloc[:, month_cols['actual']].sum()
            
            # Calculate total quantity and divide by 2 for CD and Advance CD
            total_quantity = filtered_df.iloc[:, month_cols['quantity']].sum()
            combined_data['quantity'] = total_quantity / 2  # Divide summed quantity by 2
    
     return combined_data
def main():
    processor = DiscountAnalytics()
    
    # Enhanced Sidebar
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 1rem;'>
            <h2 style='color: #1e293b;'>Dashboard Controls</h2>
        </div>
        """, unsafe_allow_html=True)
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
    
    # Enhanced Header
    st.markdown("""
    <div class='dashboard-header'>
        <div class='dashboard-title'>Discount Analytics Dashboard</div>
        <div class='dashboard-subtitle'>Monitor and analyze discount performance across states</div>
    </div>
    """, unsafe_allow_html=True)
    
    if uploaded_file is not None:
        with st.spinner('Processing data...'):
            data = processor.process_excel(uploaded_file)
            processor.create_ticker(data)
        
        # Enhanced Metrics Layout
        st.markdown("""
        <div style='margin: 2rem 0;'>
            <h3 style='color: #1e293b; margin-bottom: 1rem;'>Key Performance Indicators</h3>
        </div>
        """, unsafe_allow_html=True)
        
        processor.create_summary_metrics(data)
        
        # Enhanced Selection Controls
        st.markdown("""
        <div style='margin: 2rem 0;'>
            <h3 style='color: #1e293b; margin-bottom: 1rem;'>Detailed Analysis</h3>
        </div>
        """, unsafe_allow_html=True)
        
        col1, col2 = st.columns(2)
        with col1:
            selected_state = st.selectbox("Select State", list(data.keys()))
        
        if selected_state:
            with col2:
                discount_types = processor.get_discount_types(data[selected_state], selected_state)
                selected_discount = st.selectbox("Select Discount Type", discount_types)
            
            # Wrap charts in custom containers
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            processor.create_monthly_metrics(data, selected_state, selected_discount)
            st.markdown("</div>", unsafe_allow_html=True)
            
            st.markdown("<div class='chart-container'>", unsafe_allow_html=True)
            processor.create_trend_chart(data, selected_state, selected_discount)
            st.markdown("</div>", unsafe_allow_html=True)
    
    else:
        st.markdown("""
        <div style='text-align: center; padding: 3rem; background: white; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.1);'>
            <img src="https://via.placeholder.com/100x100" style="margin-bottom: 1rem;">
            <h2 style='color: #1e293b; margin-bottom: 1rem;'>Welcome to Discount Analytics</h2>
            <p style='color: #64748b; margin-bottom: 2rem;'>Please upload an Excel file to begin your analysis.</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Enhanced placeholder metrics
        st.markdown("<div style='margin-top: 2rem;'>", unsafe_allow_html=True)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total States", "0", "Waiting")
        with col2:
            st.metric("Total Discount Types", "0", "Waiting")
        with col3:
            st.metric("Average Discount Rate", "₹0.00", "Waiting")
        st.markdown("</div>", unsafe_allow_html=True)

if __name__ == "__main__":
    main()
