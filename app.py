import streamlit as st
import pandas as pd
import os
import glob

# Title of the dashboard
st.title("Sales Forecast Dashboard")

# Load Excel files from the templates folder
data_files = glob.glob('templates/*.xlsx')
dataframes = {}

for file in data_files:
    df_name = os.path.basename(file).split('.')[0]
    dataframes[df_name] = pd.read_excel(file)

# Display data and visualizations
for name, df in dataframes.items():
    st.subheader(f"{name} Data")
    st.dataframe(df)

    # You can add your KPI calculations and charts here
    # Example: st.metric(label="KPI Name", value="Metric Value")

    # Insert any charts that you want to visualize
    # Example: st.line_chart(data=df['some_column'])

# Additional design features and interactivity can be added as per requirements
