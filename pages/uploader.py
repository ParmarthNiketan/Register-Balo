import streamlit as st
import pandas as pd
import requests

# Streamlit app
st.title('Excel to Google Drive Uploader')

uploaded_file = st.file_uploader("Choose an Excel file", type="xlsx")

if uploaded_file is not None:
    # Read the Excel file
    df = pd.read_excel(uploaded_file)
    st.write("File content:")
    st.write(df)
    
    # Assume we have an upload endpoint
    upload_url = "https://drive.google.com/drive/u/0/home"  # Replace with actual upload URL
    
    if st.button("Upload to Google Drive"):
        files = {'file': (uploaded_file.name, uploaded_file.getvalue())}
        response = requests.post(upload_url, files=files)
        
        if response.status_code == 200:
            st.success("File uploaded successfully!")
            st.write(response.json())  # Assuming the response contains some JSON data
        else:
            st.error(f"Failed to upload file: {response.text}")