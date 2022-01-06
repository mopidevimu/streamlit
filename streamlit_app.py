# Loading Python Packages
import streamlit as st
import pandas as pd
from openpyxl import load_workbook


# Header of application
html_header="""
  <head>
    <meta charset="utf-8">
    <title>XLSX</title>
    <meta charset="utf-8">
    <meta name="keywords" content="Python xlsx comparison, XLSX comparison, Xlxs comparison">
    <meta name="description" content="comparion of tw2 excel xlsx files using python and streamlit app">
    <meta name="author" content="Murali Krishna MOPIDEVI">
    <meta name="viewport" content="width=device-width, initial-scale=1">
</head>
<center><h1 style ="">XLSX Files Comparison</h1> </center> <br>
"""
st.set_page_config(page_title="AFD", page_icon="", layout="wide")
st.markdown(html_header, unsafe_allow_html=True)

with st.container():
    xlsx_1, xlsx_2 = st.columns(2)
    with xlsx_1:
        uploaded_file_1 = st.file_uploader("Choose a Excel file 1 (XLSX)")
        if uploaded_file_1 is not None:
            wb1 = load_workbook(filename = uploaded_file_1)
            st.write(wb1)
            res = len(wb1.sheetnames)
            st.write("Number of sheets in this Xlsx1 files is: ",res)
            


    with xlsx_2:
        uploaded_file_2 = st.file_uploader("Choose a Excel file 2 (XLSX)")
        if uploaded_file_2 is not None:
            wb2 = load_workbook(filename = uploaded_file_2)
            st.write(wb2)
            res2 = len(wb2.sheetnames)
            st.write("Number of sheets in this Xlsx2 files is: ",res2)


