import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime

# Set page config
st.set_page_config(
    page_title="Sales Analytics Dashboard",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .metric-container {
        background-color: #f0f2f6;
        padding: 1rem;
        border-radius: 10px;
        margin: 0.5rem 0;
    }
    .info-box {
        background-color: #e1f5fe;
        padding: 1rem;
        border-radius: 5px;
        border-left: 5px solid #01579b;
    }
</style>
""", unsafe_allow_html=True)

def load_data(uploaded_file):
    """Load and process the Excel file"""
    try:
        df = pd.read_excel(uploaded_file)
        return df
    except Exception as e:
        st.error(f"Error loading file: {str(e)}")
        return None

def prepare_monthly_columns(df):
    """Identify monthly sales columns"""
    monthly_cols = []
    fy_cols = []
    
    # Current FY columns (FY26)
    current_fy_monthly = ["Sep'25(Till 24th)", "Aug'25", "July'25", "June'25", "May'25", "Apr'25"]
    
    # Previous FY columns (FY25)
    prev_fy_monthly = ["Mar'25", "Feb'25", "Jan'25", "Dec'24", "Nov'24", "Oct'24", 
                       "Sep'24", "Aug'24", "Jul'24", "Jun'24", "May'24", "Apr'24"]
    
    # FY24 columns
    fy24_monthly = ["Mar'24", "Feb'24", "Jan'24", "Dec'23", "Nov'23", "Oct'23",
                    "Sep'23", "Aug'23", "Jul'23", "Jun'23", "May'23", "Apr'23",
                    "Mar'23", "Feb'23", "Jan'23"]
    
    # Total columns
    total_cols = ["YTD FY26", "FY 25 Total", "FY24 Total"]
    
    return current_fy_monthly, prev_fy_monthly, fy24_monthly, total_cols

def calculate_growth(current_val, previous_val):
    """Calculate growth percentage"""
    if previous_val == 0 or pd.isna(previous_val):
        return 0
    return ((current_val - previous_val) / previous_val) * 100

def create_quarterly_data(df, monthly_cols):
    """Create quarterly aggregation from monthly data"""
    quarterly_data = {}
    
    # Define quarters (example for FY format)
    quarters = {
        'Q1': monthly_cols[9:12] if len(monthly_cols) >= 12 else [],  # Apr-Jun
        'Q2': monthly_cols[6:9] if len(monthly_cols) >= 9 else [],    # Jul-Sep
        'Q3': monthly_cols[3:6] if len(monthly_cols) >= 6 else [],    # Oct-Dec
        'Q4': monthly_cols[0:3] if len(monthly_cols) >= 3 else []     # Jan-Mar
    }
    
    for quarter, cols in quarters.items():
        if cols:
            existing_cols = [col for col in cols if col in df.columns]
            if existing_cols:
                quarterly_data[quarter] = df[existing_cols].sum(axis=1)
    
    return quarterly_data

def main():
    st.markdown('<h1 class="main-header">📊 Sales Analytics Dashboard</h1>', unsafe_allow_html=True)
    
    # Sidebar for file upload
    st.sidebar.header("📁 Data Upload")
    uploaded_file = st.sidebar.file_uploader(
        "Upload Excel File", 
        type=['xlsx', 'xls'],
        help="Upload your sales data Excel file"
    )
    
    if uploaded_file is None:
        st.info("👆 Please upload an Excel file to get started")
        st.markdown("""
        ### Expected Columns:
        - DISTRICT NAME, Region Name, ZONE NAME
        - CUSTOMER CODE, CUSTOMER NAME
        - MATERIAL, MATERIAL NAME, MATERIAL_TYPE_4
        - Monthly sales columns (Sep'25, Aug'25, etc.)
        - FY totals (YTD FY26, FY 25 Total, FY24 Total)
        """)
        return
    
    # Load data
    df = load_data(uploaded_file)
    if df is None:
        return
    
    st.success(f"✅ Data loaded successfully! ({len(df)} records)")
    
    # Data preparation
    current_fy_monthly, prev_fy_monthly, fy24_monthly, total_cols = prepare_monthly_columns(df)
    
    # Sidebar filters
    st.sidebar.header("🔍 Filters")
    
    # District selection
    districts = sorted(df['DISTRICT NAME'].dropna().unique())
    selected_district = st.sidebar.selectbox("Select District", districts)
    
    # Filter data by district
    district_data = df[df['DISTRICT NAME'] == selected_district].copy()
    
    # Display district information
    if not district_data.empty:
        district_info = district_data.iloc[0]
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown(f"**🏢 District:** {selected_district}")
            st.markdown(f"**📍 Region:** {district_info.get('Region Name', 'N/A')}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown(f"**🏷️ District Category:** {district_info.get('New District Category', 'N/A')}")
            st.markdown(f"**🌏 Zone:** {district_info.get('ZONE NAME', 'N/A')}")
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col3:
            st.markdown('<div class="info-box">', unsafe_allow_html=True)
            st.markdown(f"**👥 Total Customers:** {district_data['CUSTOMER NAME'].nunique()}")
            st.markdown(f"**📦 Total Materials:** {district_data['MATERIAL'].nunique()}")
            st.markdown('</div>', unsafe_allow_html=True)
    
    # Additional filters
    customers = sorted(district_data['CUSTOMER NAME'].dropna().unique())
    selected_customers = st.sidebar.multiselect("Select Customers (Optional)", customers)
    
    materials = sorted(district_data['MATERIAL NAME'].dropna().unique())
    selected_materials = st.sidebar.multiselect("Select Materials (Optional)", materials)
    
    # Apply additional filters
    filtered_data = district_data.copy()
    if selected_customers:
        filtered_data = filtered_data[filtered_data['CUSTOMER NAME'].isin(selected_customers)]
    if selected_materials:
        filtered_data = filtered_data[filtered_data['MATERIAL NAME'].isin(selected_materials)]
    
    # Main dashboard
    st.header("📈 Sales Overview")
    
    # Key metrics
    col1, col2, col3, col4 = st.columns(4)
    
    # Calculate key metrics
    ytd_fy26 = filtered_data['YTD FY26'].sum() if 'YTD FY26' in filtered_data.columns else 0
    fy25_total = filtered_data['FY 25 Total'].sum() if 'FY 25 Total' in filtered_data.columns else 0
    fy24_total = filtered_data['FY24 Total'].sum() if 'FY24 Total' in filtered_data.columns else 0
    
    # Growth calculations
    ytd_growth = calculate_growth(ytd_fy26, fy25_total)
    fy25_growth = calculate_growth(fy25_total, fy24_total)
    
    with col1:
        st.metric(
            label="YTD FY26 Sales",
            value=f"₹{ytd_fy26:,.0f}",
            delta=f"{ytd_growth:.1f}%" if ytd_growth != 0 else None
        )
    
    with col2:
        st.metric(
            label="FY25 Total Sales",
            value=f"₹{fy25_total:,.0f}",
            delta=f"{fy25_growth:.1f}%" if fy25_growth != 0 else None
        )
    
    with col3:
        st.metric(
            label="FY24 Total Sales", 
            value=f"₹{fy24_total:,.0f}"
        )
    
    with col4:
        avg_monthly = fy25_total / 12 if fy25_total > 0 else 0
        st.metric(
            label="Avg Monthly (FY25)",
            value=f"₹{avg_monthly:,.0f}"
        )
    
    # Tabs for different views
    tab1, tab2, tab3, tab4 = st.tabs(["📅 Monthly Trends", "🏢 Customer Analysis", "📦 Material Analysis", "📊 Growth Analysis"])
    
    with tab1:
        st.subheader("Monthly Sales Trends")
        
        # Monthly trend chart
        monthly_data = []
        labels = []
        
        # Combine current and previous FY monthly data
        all_monthly_cols = current_fy_monthly + prev_fy_monthly
        for col in all_monthly_cols:
            if col in filtered_data.columns:
                monthly_data.append(filtered_data[col].sum())
                labels.append(col)
        
        if monthly_data:
            fig = go.Figure()
            fig.add_trace(go.Scatter(
                x=labels,
                y=monthly_data,
                mode='lines+markers',
                name='Monthly Sales',
                line=dict(color='#1f77b4', width=3),
                marker=dict(size=8)
            ))
            
            fig.update_layout(
                title="Monthly Sales Trend",
                xaxis_title="Month",
                yaxis_title="Sales (₹)",
                hovermode='x unified',
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
        
        # Quarterly analysis
        st.subheader("Quarterly Performance")
        quarterly_data = create_quarterly_data(filtered_data, prev_fy_monthly)
        
        if quarterly_data:
            quarters = list(quarterly_data.keys())
            quarterly_values = [quarterly_data[q].sum() for q in quarters]
            
            col1, col2 = st.columns(2)
            
            with col1:
                fig_bar = px.bar(
                    x=quarters,
                    y=quarterly_values,
                    title="Quarterly Sales Comparison",
                    color=quarterly_values,
                    color_continuous_scale="Blues"
                )
                fig_bar.update_layout(height=300)
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with col2:
                fig_pie = px.pie(
                    values=quarterly_values,
                    names=quarters,
                    title="Quarterly Sales Distribution"
                )
                fig_pie.update_layout(height=300)
                st.plotly_chart(fig_pie, use_container_width=True)
    
    with tab2:
        st.subheader("Customer-wise Analysis")
        
        # Top customers by YTD sales
        customer_sales = filtered_data.groupby('CUSTOMER NAME').agg({
            'YTD FY26': 'sum',
            'FY 25 Total': 'sum',
            'FY24 Total': 'sum'
        }).sort_values('YTD FY26', ascending=False)
        
        customer_sales['Growth %'] = customer_sales.apply(
            lambda row: calculate_growth(row['YTD FY26'], row['FY 25 Total']), axis=1
        )
        
        # Top 10 customers chart
        top_customers = customer_sales.head(10)
        
        fig = px.bar(
            x=top_customers['YTD FY26'],
            y=top_customers.index,
            orientation='h',
            title="Top 10 Customers by YTD FY26 Sales",
            color=top_customers['Growth %'],
            color_continuous_scale="RdYlGn"
        )
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)
        
        # Customer performance table
        st.subheader("Customer Performance Summary")
        st.dataframe(customer_sales.round(2), use_container_width=True)
    
    with tab3:
        st.subheader("Material-wise Analysis")
        
        # Material analysis
        material_sales = filtered_data.groupby('MATERIAL NAME').agg({
            'YTD FY26': 'sum',
            'FY 25 Total': 'sum',
            'FY24 Total': 'sum'
        }).sort_values('YTD FY26', ascending=False)
        
        material_sales['Growth %'] = material_sales.apply(
            lambda row: calculate_growth(row['YTD FY26'], row['FY 25 Total']), axis=1
        )
        
        # Top materials chart
        top_materials = material_sales.head(10)
        
        fig = px.treemap(
            names=top_materials.index,
            values=top_materials['YTD FY26'],
            title="Top Materials by Sales Volume (YTD FY26)"
        )
        fig.update_layout(height=500)
        st.plotly_chart(fig, use_container_width=True)
        
        # Material type analysis
        if 'MATERIAL_TYPE_4' in filtered_data.columns:
            material_type_sales = filtered_data.groupby('MATERIAL_TYPE_4')['YTD FY26'].sum().sort_values(ascending=False)
            
            fig = px.pie(
                values=material_type_sales.values,
                names=material_type_sales.index,
                title="Sales by Material Type"
            )
            st.plotly_chart(fig, use_container_width=True)
    
    with tab4:
        st.subheader("Growth Analysis")
        
        # Growth comparison chart
        growth_data = {
            'Period': ['YTD FY26 vs FY25', 'FY25 vs FY24'],
            'Growth %': [ytd_growth, fy25_growth],
            'Current': [ytd_fy26, fy25_total],
            'Previous': [fy25_total, fy24_total]
        }
        
        growth_df = pd.DataFrame(growth_data)
        
        fig = px.bar(
            growth_df,
            x='Period',
            y='Growth %',
            title="Year-over-Year Growth Comparison",
            color='Growth %',
            color_continuous_scale="RdYlGn"
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Monthly growth analysis
        st.subheader("Monthly Growth Trends")
        
        # Compare current year months with previous year
        monthly_growth = []
        growth_labels = []
        
        current_months = ["Sep'25(Till 24th)", "Aug'25", "July'25", "June'25", "May'25", "Apr'25"]
        previous_months = ["Sep'24", "Aug'24", "Jul'24", "Jun'24", "May'24", "Apr'24"]
        
        for curr, prev in zip(current_months, previous_months):
            if curr in filtered_data.columns and prev in filtered_data.columns:
                curr_val = filtered_data[curr].sum()
                prev_val = filtered_data[prev].sum()
                growth = calculate_growth(curr_val, prev_val)
                monthly_growth.append(growth)
                growth_labels.append(curr.replace("'25", "").replace("(Till 24th)", ""))
        
        if monthly_growth:
            fig = go.Figure()
            colors = ['green' if x >= 0 else 'red' for x in monthly_growth]
            
            fig.add_trace(go.Bar(
                x=growth_labels,
                y=monthly_growth,
                marker_color=colors,
                text=[f"{x:.1f}%" for x in monthly_growth],
                textposition='outside'
            ))
            
            fig.update_layout(
                title="Month-over-Month Growth (FY26 vs FY25)",
                xaxis_title="Month",
                yaxis_title="Growth %",
                height=400
            )
            
            st.plotly_chart(fig, use_container_width=True)
    
    # Data export section
    st.header("📥 Export Data")
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("Download Filtered Data as CSV"):
            csv = filtered_data.to_csv(index=False)
            st.download_button(
                label="📁 Download CSV",
                data=csv,
                file_name=f"sales_data_{selected_district}_{datetime.now().strftime('%Y%m%d')}.csv",
                mime="text/csv"
            )
    
    with col2:
        if st.button("Download Summary Report"):
            # Create summary report
            summary = f"""
            Sales Dashboard Summary Report
            Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M')}
            
            District: {selected_district}
            Region: {district_info.get('Region Name', 'N/A')}
            Zone: {district_info.get('ZONE NAME', 'N/A')}
            
            Key Metrics:
            - YTD FY26 Sales: ₹{ytd_fy26:,.0f}
            - FY25 Total Sales: ₹{fy25_total:,.0f}
            - YTD Growth: {ytd_growth:.1f}%
            - Total Customers: {filtered_data['CUSTOMER NAME'].nunique()}
            - Total Materials: {filtered_data['MATERIAL'].nunique()}
            """
            
            st.download_button(
                label="📋 Download Summary",
                data=summary,
                file_name=f"summary_report_{selected_district}_{datetime.now().strftime('%Y%m%d')}.txt",
                mime="text/plain"
            )

if __name__ == "__main__":
    main()
