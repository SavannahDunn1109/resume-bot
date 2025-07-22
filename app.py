import streamlit as st
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import io
import pandas as pd
import os
from PyPDF2 import PdfReader
from docx import Document

# ========== CONFIG ==========
SITE_URL = "https://eleven090.sharepoint.com/sites/Recruiting"
LIBRARY = "Shared Documents"
FOLDER = "Active Resumes"

# ========== AUTH ==========
@st.cache_resource
def connect_to_sharepoint():
    ctx_auth = AuthenticationContext(SITE_URL)
    if not ctx_auth.acquire_token_for_user(
        st.secrets["sharepoint"]["username"],
        st.secrets["sharepoint"]["password"]
    ):
        st.error("Authentication failed")
        return None
    return ClientContext(SITE_URL, ctx_auth)

# ========== FILE HELPERS ==========
def download_file(ctx, file_url):
    response = File.open_binary(ctx, file_url)
    return io.BytesIO(response.content)

def extract_text_from_pdf(file_bytes):
    text = ""
    reader = PdfReader(file_bytes)
    for page in reader.pages:
        page_text = page.extract_text()
        if page_text:
            text += page_text + "\n"
    return text

def extract_text_from_docx(file_bytes):
    doc = Document(file_bytes)
    return "\n".join([para.text for para in doc.paragraphs])

# ========== SCORING ==========
KEYWORDS = [
    "RMF", "STIGs", "NIST 800-53", "NIST 800-37", "Security Authorization", "ATO",
    "POA&M", "IAVM", "DIACAP", "eMASS", "ACAS", "HBSS",
    "Security+ CE", "CISSP", "CASP+", "CEH", "DoD 8570", "IAT II", "IAT III",
    "System Security Engineering", "Secure Architecture Design", "Vulnerability Assessment",
    "Incident Response Planning", "Audit & Compliance", "Cross-functional Team Collaboration",
    "Documentation & Risk Analysis", "Security Control Assessment",
    "Splunk", "Nessus", "Tenable", "Wireshark", "Snort", "Palo Alto"
]

def score_resume(text):
    score = 0
    found_keywords = []
    for kw in KEYWORDS:
        if kw.lower() in text.lower():
            score += 10
            found_keywords.append(kw)
    return score, ", ".join(found_keywords)

# ========== MAIN STREAMLIT APP ==========
st.title("📄 Resume Scorer from SharePoint")

ctx = connect_to_sharepoint()
if ctx:
    folder_url = f"{LIBRARY}/{FOLDER}"
    folder = ctx.web.get_folder_by_server_relative_url(folder_url)
    files = folder.files
    ctx.load(files)
    ctx.execute_query()

    data = []

    for file in files:
        filename = file.properties["Name"]
        if filename.lower().endswith(".pdf") or filename.lower().endswith(".docx"):
            file_url = file.properties["ServerRelativeUrl"]
            file_bytes = download_file(ctx, file_url)

            if filename.lower().endswith(".pdf"):
                text = extract_text_from_pdf(file_bytes)
            else:
                text = extract_text_from_docx(file_bytes)

            score, keywords_found = score_resume(text)
            data.append({
                "File Name": filename,
                "Score": score,
                "Keywords Found": keywords_found
            })

    df = pd.DataFrame(data)
    st.dataframe(df)

    if not df.empty:
        output = io.BytesIO()
        df.to_excel(output, index=False)
        output.seek(0)
        st.download_button("📥 Download Excel Report", output, file_name="resume_scores.xlsx")

        # Upload Excel back to SharePoint
        if st.button("📤 Upload Excel to SharePoint"):
            upload_path = f"{LIBRARY}/resume_scores.xlsx"
            target_folder = ctx.web.get_folder_by_server_relative_url(LIBRARY)
            target_folder.upload_file("resume_scores.xlsx", output)
            ctx.execute_query()
            st.success("Excel uploaded to SharePoint!")

 
