import pandas as pd
from ydata_profiling import ProfileReport
import streamlit as st


st.set_page_config(
    page_title="Data Profiler",
    layout="centered",
    initial_sidebar_state="expanded",
)

hide_streamlit_style = """
            <style>
            footer {visibility: hidden;}
            </style>
            """
st.markdown(hide_streamlit_style, unsafe_allow_html=True)

st.markdown("# Welcome to Data Profiler!")

uploaded_file = st.file_uploader("Upload a csv file", type=['csv'])
if uploaded_file is not None:
    try:
        file = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error(f"Error reading file: {str(e)}")
    else:
        with st.spinner('Generating the report for download...'):
            report = ProfileReport(file).to_html()
            st.download_button('Download', report, 'report.html')
