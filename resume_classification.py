import streamlit as st
import pandas as pd
import pickle
import PyPDF2
import docx2txt
import pythoncom
from win32com import client
import os

# Streamlit UI
st.title('Resume Classification and Skill Matching')

# Load the pre-trained SVC model and DataFrame
svc_model = pickle.load(open(r"C:\Users\Dell\PycharmProjects\Resume Prediction\resume_svm_model.pkl", 'rb'))
df = pickle.load(open(r"C:\Users\Dell\PycharmProjects\Resume Prediction\dataframe.pkl", 'rb'))

# Input files (resumes)
uploaded_files = st.file_uploader("Upload your resumes", type=['pdf', 'doc', 'docx'], accept_multiple_files=True)
skills = st.multiselect("Select your skills:", ["Python", "JavaScript", "Java", "C++", "SQL", "C", "Embedded C"])


# Function to extract text from a PDF file
def extract_text_from_pdf(pdf_file):
    text = ""
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text


# Function to extract text from a .docx file
def extract_text_from_docx(docx_file):
    text = docx2txt.process(docx_file)
    return text


# Function to extract text from a .doc file using win32com
def extract_text_from_doc(doc_file):
    # Save the uploaded file temporarily
    with open("temp_doc_file.doc", "wb") as f:
        f.write(doc_file.getbuffer())

    pythoncom.CoInitialize()  # Initialize the COM library
    word = client.Dispatch("Word.Application")

    # Open the temporarily saved file using win32com
    doc = word.Documents.Open(os.path.abspath("temp_doc_file.doc"))
    doc_text = doc.Range().Text
    doc.Close()
    word.Quit()

    # Remove the temporary file after processing
    os.remove("temp_doc_file.doc")

    return doc_text


# Process the uploaded files
if uploaded_files and skills:
    for uploaded_file in uploaded_files:
        file_extension = uploaded_file.name.split('.')[-1]

        # Extract text based on file type
        if file_extension == 'pdf':
            resume_text = extract_text_from_pdf(uploaded_file)
        elif file_extension == 'docx':
            resume_text = extract_text_from_docx(uploaded_file)
        elif file_extension == 'doc':
            resume_text = extract_text_from_doc(uploaded_file)
        else:
            st.write(f"Unsupported file format: {file_extension}")
            continue

        # Display the file name
        st.write(f"### Resume: {uploaded_file.name}")

        # Find matching skills
        matched_skills = [skill for skill in skills if skill.lower() in resume_text.lower()]
        if matched_skills:
            st.write(f"**Matched Skills:** {', '.join(matched_skills)}")
        else:
            st.write("No skills matched.")
