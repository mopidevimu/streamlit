# Loading Python Packages
import streamlit as st

# Header of application
html_header="""
<head>
<title>XLSX</title>
<meta charset="utf-8">
<meta name="keywords" content="Python xlsx comparison, XLSX comparison, Xlxs comparison">
<meta name="description" content="comparion of tw2 excel xlsx files using python and streamlit app">
<meta name="author" content="Murali Krishna MOPIDEVI">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
</head>
<center><h1 style ="">XLSX Files Comparison</h1> </center> <br>
"""
st.set_page_config(page_title="AFD", page_icon="", layout="wide")
st.markdown('<style>body{background-color: #fbfff0}</style>',unsafe_allow_html=True)
st.markdown(html_header, unsafe_allow_html=True)
