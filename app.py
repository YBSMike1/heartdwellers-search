import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
import tempfile

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# Light grey page background
st.markdown("""
<style>
    .stApp, .main {
        background-color: #f0f0f0 !important;
    }
    /* Folder headers */
    div[data-testid="stExpander"] > div > div > div > div > button {
        background-color: #f8e8f0 !important;
        border: 3px solid #D81B60 !important;
    }
    /* Result boxes */
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #FFCCE0 !important;
        border-left: 6px solid #D81B60 !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

DOCX_FOLDER = "Heartdwellers Docxs"

def search_italic_text(search_word, folder_path):
    results = []
    file_count = 0
    match_count = 0
    pattern = re.compile(rf'(?<!\w){re.escape(search_word)}(?!\w)', re.IGNORECASE)

    if not os.path.exists(folder_path):
        st.error(f"Folder not found: {folder_path}")
        return [], 0, 0

    progress_bar = st.progress(0)
    status = st.empty()

    total_files = sum(1 for _, _, files in os.walk(folder_path) 
                     for f in files if f.lower().endswith('.docx'))

    processed = 0
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if not filename.lower().endswith('.docx'):
                continue
            if any(skip in filename.lower() for skip in ["compilation ", "~$", "eom", "all messages"]):
                continue

            file_path = os.path.join(root, filename)
            file_count += 1
            processed += 1
            progress_bar.progress(min(processed / max(total_files, 1), 1.0))
            status.text(f"Searching: {filename}")

            try:
                doc = Document(file_path)
                for p in doc.paragraphs:
                    italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
                    if italic_text and re.search(pattern, italic_text):
                        italic_text = italic_text.strip()
                        if italic_text:
                            results.append({
                                "file": os.path.relpath(file_path, folder_path),
                                "text": italic_text
                            })
                            match_count += 1
            except:
                continue

    progress_bar.progress(1.0)
    return results, file_count, match_count

search_word = st.text_input("Enter the word or phrase to search:", placeholder="e.g. rapture, love, faith")

if st.button("🔍
