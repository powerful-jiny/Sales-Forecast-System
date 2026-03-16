import streamlit as st
import pandas as pd

# Title of the app
st.title('Sales Forecasting Application')

# File uploader for Account Master data
st.header('Upload Account Master Data')
account_master_file = st.file_uploader('Choose a CSV file', type='csv')

if account_master_file is not None:
    account_master_data = pd.read_csv(account_master_file)
    st.write(account_master_data)

# File uploader for Annual Plan data
st.header('Upload Annual Plan Data')
annual_plan_file = st.file_uploader('Choose a CSV file', type='csv', key='annual_plan')

if annual_plan_file is not None:
    annual_plan_data = pd.read_csv(annual_plan_file)
    st.write(annual_plan_data)

# File uploader for Forecast data
st.header('Upload Forecast Data')
forecast_file = st.file_uploader('Choose a CSV file', type='csv', key='forecast')

# File uploader for Weekly Actual data
st.header('Upload Weekly Actual Data')
weekly_actual_file = st.file_uploader('Choose a CSV file', type='csv', key='weekly_actual')
