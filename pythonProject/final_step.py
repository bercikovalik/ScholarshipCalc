import streamlit as st


def main():
    st.title("Final step")
    st.subheader("Upload original file with all students")
    uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")
