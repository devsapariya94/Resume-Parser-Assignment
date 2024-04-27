import streamlit as st
import pandas as pd
import re
from io import BytesIO
import textract
import os
import time

st.title("Resume Parser")
st.write("Upload your resume(s) to extract emails and phone numbers")
files = st.file_uploader("Upload your files", accept_multiple_files=True)

if files:
    # make dir ouput if it doesn't exist
    if not os.path.exists("output"):
        os.makedirs("output")

    # remove all files which is 10 minutes old
    for file in os.listdir("output"):
        if os.path.getmtime("output/"+file) < time.time() - 600:
            os.remove("output/"+file)

    # save the files to the server
    for file in files:
        with open("output/"+file.name, "wb") as f:
            f.write(file.getbuffer())

    # Extract data
    with st.spinner("Extracting data..."):
        all_data = []
        for file in files:
            text = textract.process("output/"+file.name)
            all_data.append(text.decode('utf-8'))
    
        all_text = []
        all_emails = []
        all_phone_numbers = []

        for text in all_data:
            # Extract emails
            emails = re.findall(r"[a-zA-Z0-9\.\-+]+@[a-zA-Z0-9\.\-+]+\.[a-zA-Z]+", text)
            # Extract phone numbers
            phone_numbers = re.findall(r"\b\d{10}\b", text)

            all_text.append(text)  
            all_emails.append(emails)
            all_phone_numbers.append(phone_numbers)

        # Create DataFrame
        df = pd.DataFrame({
            'Emails': all_emails,
            'Phone Numbers': all_phone_numbers,
            'Text': all_text
        })


        # Apply encoding and decoding check to DataFrame entries
        df = df.applymap(lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

       
        # Save DataFrame to BytesIO object  
        excel_data = BytesIO()
        df.to_excel(excel_data, index=False)

        # Download button
        st.download_button(label="Download", data=excel_data, file_name="output.xlsx", mime="application/vnd.ms-excel")