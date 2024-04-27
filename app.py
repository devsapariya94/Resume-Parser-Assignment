import streamlit as st
from PyPDF2 import PdfReader
import pandas as pd
import re
from io import BytesIO
from docx import Document  # Importing Document class from python-docx
import os
import win32com.client

st.title("Resume Parser")
st.write("Upload your resume(s) to extract emails and phone numbers")
files = st.file_uploader("Upload your files", accept_multiple_files=True)

# Create a lock for handling Word instances
word_lock = None
word = None

def initialize_word():
    global word
    try:
        word = win32com.client.Dispatch("Word.Application")
    except Exception as e:
        st.error(f"Error initializing Word: {e}")

def process_doc_file(file):
    global word_lock, word
    data = ""
    try:
        if not word_lock:
            word_lock = True
            initialize_word()
        # Create a temporary file to work around the issue with win32com
        with open("temp.doc", "wb") as temp_doc:
            temp_doc.write(file.read())
        
        # Get the full path to the temporary file
        temp_file_path = os.path.abspath("temp.doc")

        # Open the document
        doc = word.Documents.Open(temp_file_path)

        # Access the content of the document
        data = doc.Content.Text

        # Close the document
        doc.Close()

    except Exception as e:
        st.error(f"Error processing file: {e}")

    finally:
        # Delete the temporary file
        os.remove(temp_file_path)

    return data

def remove_illegal_characters(text):
    # Define a regular expression pattern to match illegal characters
    illegal_pattern = re.compile(r'[\\/*?[\]:]')

    # Replace illegal characters with an empty string
    sanitized_text = illegal_pattern.sub('', text)

    return sanitized_text

if files:
    with st.spinner("Extracting data..."):
        all_data = []
        for file in files:
            if file.name.endswith(".pdf"):
                pdf_reader = PdfReader(file)
                text = ""
                for page in pdf_reader.pages:
                    text += page.extract_text()
                all_data.append(text)

            elif file.name.endswith(".docx"):
                # Handling .docx files
                docx_file = Document(file)  # Using Document class from python-docx
                text = '\n'.join([paragraph.text for paragraph in docx_file.paragraphs])
                all_data.append(text)

            elif file.name.endswith(".doc"):  # Handling .doc files with win32com
                data = process_doc_file(file)
                all_data.append(data)

            else:
                all_data.append("File type not supported")

        # Extracted text, emails, and phone numbers lists
        all_text = []
        all_emails = []
        all_phone_numbers = []

        for text in all_data:
            # Extract emails
            emails = re.findall(r"[a-zA-Z0-9\.\-+]+@[a-zA-Z0-9\.\-+]+\.[a-zA-Z]+", text)
            # Extract phone numbers
            phone_numbers = re.findall(r"\b\d{10}\b", text)

            all_text.append(remove_illegal_characters(text))  # Remove illegal characters from text
            all_emails.append(emails)
            all_phone_numbers.append(phone_numbers)

        # Create DataFrame
        df = pd.DataFrame({
            'Emails': all_emails,
            'Phone Numbers': all_phone_numbers,
            'Text': all_text
        })

        # Apply the check to encode and decode strings
        df = df.applymap(lambda x: x.encode('unicode_escape').decode('utf-8') if isinstance(x, str) else x)

        # Save DataFrame to BytesIO object  
        excel_data = BytesIO()
        df.to_excel(excel_data, index=False)

        # Download button
        st.download_button(label="Download", data=excel_data, file_name="output.xlsx", mime="application/vnd.ms-excel")


