import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import distinctipy
from pathlib import Path
import io

def load_and_process_data(uploaded_file):
    """Load Excel file and return dict of dataframes and sheet names"""
    xl = pd.ExcelFile(uploaded_file)
    states = xl.sheet_names
    state_dfs = {state: pd.read_excel(uploaded_file, sheet_name=state) for state in states}
    return state_dfs, states

def get_available_months(df):
    """Extract available months from column names"""
    share_cols = [col for col in df.columns if col.startswith('Share_')]
    months = [col.split('_')[1] for col in share_cols]
    return sorted(list(set(months)))

def create_share_plot(df, selected_months):
    """Create stacked bar chart for selected months"""
    # Process data for selected months
    all_data = []
    for month in selected_months:
        month_data = df[['Company', f'Share_{month}', f'WSP_{month}']].copy()
        month_data.columns = ['Company', 'Share', 'WSP']
        month_data['Month'] = month.capitalize()
        all_data.append(month_data)
    
    combined_df = pd.concat(all_data, ignore_index=True)
    
    # Calculate price ranges
    min_price = (combined_df['WSP'].min() // 5) * 5
    max_price = (combined_df['WSP'].max() // 5 + 1) * 5
    price_ranges = pd.interval_range(start=min_price, end=max_price, freq=5)
    
    # Create price range column
    combined_df['Price_Range'] = pd.cut(combined_df['WSP'], bins=price_ranges)
    
    # Create pivot table
    pivot_df = pd.pivot_table(
        combined_df,
        values='Share',
        index=['Month', 'Price_Range'],
        columns='Company',
        aggfunc='sum',
        fill_value=0
    )
    
    # Remove zero columns
    pivot_df = pivot_df.loc[:, (pivot_df != 0).any(axis=0)]
    
    # Calculate row sums
    row_sums = pivot_df.sum(axis=1)
    
    # Create plot
    fig, ax = plt.subplots(figsize=(15, 8))
    
    # Generate colors
    n_companies = len(pivot_df.columns)
    colors = distinctipy.get_colors(n_companies)
    
    # Plot stacked bars
    pivot_df.plot(
        kind='bar',
        stacked=True,
        ax=ax,
        width=0.8,
        color=colors
    )
    
    # Customize plot
    plt.title('Company Shares Distribution by WSP Price Range',
             fontsize=18,
             pad=20,
             fontweight='bold')
    plt.xlabel('WSP Price Range', fontsize=14)
    plt.ylabel('Share (%)', fontsize=14)
    
    # Rotate x-axis labels
    plt.xticks(rotation=45, ha='right')
    
    # Add percentage labels
    for c in ax.containers:
        labels = [f'{v:.1f}%' if v > 0 else '' for v in c.datavalues]
        ax.bar_label(c, labels=labels, label_type='center')
    
    # Add total labels
    for i, (idx, total) in enumerate(row_sums.items()):
        ax.text(i, total + 0.5, f'Total: {total:.1f}%',
                ha='center',
                va='bottom',
                fontweight='bold',
                bbox=dict(facecolor='white',
                         edgecolor='none',
                         alpha=0.7,
                         pad=3))
    
    # Customize legend
    plt.legend(
        bbox_to_anchor=(1.05, 1),
        loc='upper left',
        borderaxespad=0.,
        frameon=True,
        fontsize=12,
        title='Companies',
        title_fontsize=14
    )
    
    # Add grid
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    
    # Customize background
    ax.set_facecolor('#f8f9fa')
    fig.patch.set_facecolor('#ffffff')
    
    # Adjust layout
    current_ymax = ax.get_ylim()[1]
    ax.set_ylim(0, current_ymax * 1.1)
    plt.margins(y=0.1)
    plt.tight_layout()
    
    return fig

def main():
    # Set page config
    st.set_page_config(
        page_title="Market Share Analysis",
        page_icon="📊",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Custom CSS
    st.markdown("""
        <style>
        .main {
            background-color: #f8f9fa;
        }
        .stApp {
            background-color: #ffffff;
        }
        .css-1d391kg {
            padding: 2rem 1rem;
        }
        .stSelectbox {
            background-color: white;
        }
        .stMultiSelect {
            background-color: white;
        }
        </style>
    """, unsafe_allow_html=True)
    
    # Header
    st.title("📊 Market Share Analysis Dashboard")
    st.markdown("---")
    
    # Sidebar
    with st.sidebar:
        st.header("📥 Data Upload")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        
        if uploaded_file:
            state_dfs, states = load_and_process_data(uploaded_file)
            
            st.header("🎯 Analysis Settings")
            # State selection
            selected_state = st.selectbox(
                "Select State",
                states,
                index=0,
                help="Choose the state for analysis"
            )
            
            # Get available months for selected state
            available_months = get_available_months(state_dfs[selected_state])
            
            # Month selection
            selected_months = st.multiselect(
                "Select Months",
                available_months,
                default=[available_months[0]],
                help="Choose one or more months for comparison"
            )
            
            if not selected_months:
                st.warning("Please select at least one month.")
                return
            
            # Create and display plot
            with st.spinner("Generating visualization..."):
                fig = create_share_plot(state_dfs[selected_state], selected_months)
                
                # Main content area
                st.markdown("### 📈 Market Share Distribution")
                st.pyplot(fig)
                
                # Download button for the plot
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
                buf.seek(0)
                st.download_button(
                    label="Download Plot",
                    data=buf,
                    file_name=f'market_share_{selected_state}_{"-".join(selected_months)}.png',
                    mime='image/png'
                )
        else:
            st.info("👆 Upload an Excel file to begin analysis")
    
    # Footer
    st.markdown("---")
    st.markdown(
        """
        <div style='text-align: center'>
            <p>Built with ❤️ using Streamlit</p>
        </div>
        """,
        unsafe_allow_html=True
    )

if __name__ == "__main__":
    main()
