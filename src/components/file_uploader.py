import streamlit as st

def upload_files():
    return st.file_uploader("Vælg flere ad gangen", type=['ics', 'ical'], accept_multiple_files=True)