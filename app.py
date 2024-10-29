import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px

def main():
    st.title("Interactive Data Visualization App")
    st.write("Upload your dataset and create various plots with generated code!")

    # File upload
    uploaded_file = st.file_uploader("Choose a CSV file", type="csv")
    
    if uploaded_file is not None:
        # Read the data
        df = pd.read_csv(uploaded_file)
        
        # Show dataset preview
        st.subheader("Dataset Preview")
        st.dataframe(df.head())
        
        # Show column information
        st.subheader("Column Information")
        st.write("Numerical columns:", df.select_dtypes(include=['float64', 'int64']).columns.tolist())
        st.write("Categorical columns:", df.select_dtypes(include=['object']).columns.tolist())
        
        # Plot selection
        plot_type = st.selectbox(
            "Select the type of plot you want to create",
            ["Scatter Plot", "Line Plot", "Bar Plot", "Histogram", "Box Plot", "Heatmap", "Pie Chart"]
        )
        
        # Column selection based on plot type
        if plot_type in ["Scatter Plot"]:
            x_col = st.selectbox("Select X-axis column", df.select_dtypes(include=['float64', 'int64']).columns)
            y_col = st.selectbox("Select Y-axis column", df.select_dtypes(include=['float64', 'int64']).columns)
            color_col = st.selectbox("Select color column (optional)", ['None'] + df.columns.tolist())
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_scatter_code(x_col, y_col, color_col)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    if color_col != 'None':
                        fig = px.scatter(df, x=x_col, y=y_col, color=color_col, title=f"{x_col} vs {y_col}")
                    else:
                        fig = px.scatter(df, x=x_col, y=y_col, title=f"{x_col} vs {y_col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Line Plot":
            x_col = st.selectbox("Select X-axis column", df.columns)
            y_col = st.selectbox("Select Y-axis column", df.select_dtypes(include=['float64', 'int64']).columns)
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_line_code(x_col, y_col)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    fig = px.line(df, x=x_col, y=y_col, title=f"{y_col} over {x_col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Bar Plot":
            x_col = st.selectbox("Select X-axis column", df.columns)
            y_col = st.selectbox("Select Y-axis column", df.select_dtypes(include=['float64', 'int64']).columns)
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_bar_code(x_col, y_col)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    fig = px.bar(df, x=x_col, y=y_col, title=f"{y_col} by {x_col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Histogram":
            col = st.selectbox("Select column", df.select_dtypes(include=['float64', 'int64']).columns)
            bins = st.slider("Number of bins", min_value=5, max_value=50, value=20)
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_histogram_code(col, bins)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    fig = px.histogram(df, x=col, nbins=bins, title=f"Histogram of {col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Box Plot":
            y_col = st.selectbox("Select column for box plot", df.select_dtypes(include=['float64', 'int64']).columns)
            x_col = st.selectbox("Select grouping column (optional)", ['None'] + df.columns.tolist())
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_box_code(x_col, y_col)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    if x_col != 'None':
                        fig = px.box(df, x=x_col, y=y_col, title=f"Box Plot of {y_col} by {x_col}")
                    else:
                        fig = px.box(df, y=y_col, title=f"Box Plot of {y_col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Heatmap":
            num_cols = df.select_dtypes(include=['float64', 'int64']).columns
            selected_cols = st.multiselect("Select columns for correlation heatmap", num_cols, default=num_cols[:5])
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_heatmap_code(selected_cols)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    corr = df[selected_cols].corr()
                    fig = px.imshow(corr, 
                                  title="Correlation Heatmap",
                                  color_continuous_scale="RdBu")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")
                    
        elif plot_type == "Pie Chart":
            col = st.selectbox("Select column for pie chart", df.columns)
            
            # Generate and display code
            st.subheader("Generated Code")
            code = generate_pie_code(col)
            st.code(code, language='python')
            
            # Create and display plot
            if st.button("Generate Plot"):
                try:
                    value_counts = df[col].value_counts()
                    fig = px.pie(values=value_counts.values, 
                               names=value_counts.index, 
                               title=f"Pie Chart of {col}")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.error(f"Error generating plot: {str(e)}")

def generate_scatter_code(x_col, y_col, color_col):
    if color_col != 'None':
        return f"""import plotly.express as px

# Create scatter plot
fig = px.scatter(df, x='{x_col}', y='{y_col}', color='{color_col}',
                 title='{x_col} vs {y_col}')

# Show the plot
fig.show()"""
    else:
        return f"""import plotly.express as px

# Create scatter plot
fig = px.scatter(df, x='{x_col}', y='{y_col}',
                 title='{x_col} vs {y_col}')

# Show the plot
fig.show()"""

def generate_line_code(x_col, y_col):
    return f"""import plotly.express as px

# Create line plot
fig = px.line(df, x='{x_col}', y='{y_col}',
              title='{y_col} over {x_col}')

# Show the plot
fig.show()"""

def generate_bar_code(x_col, y_col):
    return f"""import plotly.express as px

# Create bar plot
fig = px.bar(df, x='{x_col}', y='{y_col}',
             title='{y_col} by {x_col}')

# Show the plot
fig.show()"""

def generate_histogram_code(col, bins):
    return f"""import plotly.express as px

# Create histogram
fig = px.histogram(df, x='{col}', nbins={bins},
                   title='Histogram of {col}')

# Show the plot
fig.show()"""

def generate_box_code(x_col, y_col):
    if x_col != 'None':
        return f"""import plotly.express as px

# Create box plot
fig = px.box(df, x='{x_col}', y='{y_col}',
             title='Box Plot of {y_col} by {x_col}')

# Show the plot
fig.show()"""
    else:
        return f"""import plotly.express as px

# Create box plot
fig = px.box(df, y='{y_col}',
             title='Box Plot of {y_col}')

# Show the plot
fig.show()"""

def generate_heatmap_code(selected_cols):
    cols_str = "', '".join(selected_cols)
    return f"""import plotly.express as px

# Calculate correlation matrix
corr = df[['{cols_str}']].corr()

# Create heatmap
fig = px.imshow(corr,
                title='Correlation Heatmap',
                color_continuous_scale='RdBu')

# Show the plot
fig.show()"""

def generate_pie_code(col):
    return f"""import plotly.express as px

# Calculate value counts
value_counts = df['{col}'].value_counts()

# Create pie chart
fig = px.pie(values=value_counts.values,
             names=value_counts.index,
             title='Pie Chart of {col}')

# Show the plot
fig.show()"""

if __name__ == "__main__":
    main()
