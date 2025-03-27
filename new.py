import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import distinctipy
from pathlib import Path
import io
def load_and_process_data(uploaded_file):
    xl = pd.ExcelFile(uploaded_file)
    states = xl.sheet_names
    state_dfs = {state: pd.read_excel(uploaded_file, sheet_name=state) for state in states}
    return state_dfs, states
def get_available_months(df):
    share_cols = [col for col in df.columns if col.startswith('Share_')]
    months = [col.split('_')[1] for col in share_cols]
    return sorted(list(set(months)))
def create_share_plot(df, month):
    month_data = df[['Company', f'Share_{month}', f'WSP_{month}']].copy()
    month_data.columns = ['Company', 'Share', 'WSP']
    min_price = (month_data['WSP'].min() // 5) * 5
    max_price = (month_data['WSP'].max() // 5 + 1) * 5
    price_ranges = pd.interval_range(start=min_price, end=max_price, freq=5)
    month_data['Price_Range'] = pd.cut(month_data['WSP'], bins=price_ranges)
    pivot_df = pd.pivot_table(month_data,values='Share',index='Price_Range',columns='Company',aggfunc='sum',fill_value=0)
    pivot_df = pivot_df.loc[:, (pivot_df != 0).any(axis=0)]
    row_sums = pivot_df.sum(axis=1)
    fig, ax = plt.subplots(figsize=(12, 6))
    n_companies = len(pivot_df.columns)
    colors = distinctipy.get_colors(n_companies)
    pivot_df.plot(kind='bar',stacked=True,ax=ax,width=0.8,color=colors)
    plt.title(f'Company Shares Distribution by WSP Price Range ({month.capitalize()})',fontsize=16,pad=20,fontweight='bold')
    plt.xlabel('WSP Price Range', fontsize=12)
    plt.ylabel('Share (%)', fontsize=12)
    def format_interval(interval):
        return f'{interval.left:.0f}-{interval.right:.0f}'
    x_labels = [format_interval(interval) for interval in pivot_df.index]
    ax.set_xticklabels(x_labels, rotation=0, ha='center')
    for c in ax.containers:
        labels = [f'{v:.1f}%' if v > 0 else '' for v in c.datavalues]
        ax.bar_label(c, labels=labels, label_type='center')
        for i, (idx, total) in enumerate(row_sums.items()):
        ax.text(i, total + 0.5, f'Total: {total:.1f}%',ha='center',va='bottom',fontweight='bold',bbox=dict(facecolor='white',edgecolor='none',alpha=0.7,pad=3))
    plt.legend(bbox_to_anchor=(1.05, 1),loc='upper left',borderaxespad=0.,frameon=True,fontsize=10,title='Companies',title_fontsize=12)
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    ax.set_facecolor('#f8f9fa')
    fig.patch.set_facecolor('#ffffff')
    current_ymax = ax.get_ylim()[1]
    ax.set_ylim(0, current_ymax * 1.1)
    plt.margins(y=0.1)
    plt.tight_layout()
    return fig
def main():
    st.set_page_config(page_title="Market Share Analysis",page_icon="📊",layout="wide",initial_sidebar_state="expanded")
    st.markdown("""<style>.main {background-color: #f8f9fa;}.stApp {background-color: #ffffff;}.css-1d391kg {padding: 2rem 1rem;}.stSelectbox {background-color: white;}.stMultiSelect {background-color: white;}</style>""", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 4])
    with col1:
        st.header("📥 Data Upload")
        uploaded_file = st.file_uploader("Upload Excel File", type=['xlsx', 'xls'])
        if uploaded_file:
            state_dfs, states = load_and_process_data(uploaded_file)
            st.header("🎯 Analysis Settings")
            selected_state = st.selectbox("Select State",states,index=0,help="Choose the state for analysis")
            available_months = get_available_months(state_dfs[selected_state])
            selected_months = st.multiselect("Select Months",available_months,default=[available_months[0]],help="Choose one or more months for comparison")
            if not selected_months:
                st.warning("Please select at least one month.")
                return
    with col2:
        if uploaded_file and selected_months:
            st.title("📊 Market Share Analysis Dashboard")
            st.markdown(f"### State: {selected_state}")
            st.markdown("---")
            for month in selected_months:
                with st.spinner(f"Generating visualization for {month.capitalize()}..."):
                    fig = create_share_plot(state_dfs[selected_state], month)
                    st.pyplot(fig)
                    buf = io.BytesIO()
                    fig.savefig(buf, format='png', dpi=300, bbox_inches='tight')
                    buf.seek(0)
                    st.download_button(
                        label=f"Download {month.capitalize()} Plot",
                        data=buf,
                        file_name=f'market_share_{selected_state}_{month}.png',
                        mime='image/png',
                        key=f"download_{month}"  # Unique key for each button
                    )
                st.markdown("---")
        
        elif uploaded_file:
            st.info("👈 Select state and months from the sidebar to view analysis")
        else:
            st.info("👈 Upload an Excel file from the sidebar to begin analysis")
    
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
