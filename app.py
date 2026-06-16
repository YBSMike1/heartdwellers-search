import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
import tempfile
import requests
from datetime import datetime

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
WORDS_API_KEY = "e10c87331emshb838f6dd5aeb4e8p1a63dbjsn139eec4460a9"

def get_word_definition(word):
    if not word or len(word) < 2:
        return "Please enter a valid word."
    try:
        url = f"https://wordsapiv1.p.rapidapi.com/words/{word.lower()}/definitions"
        headers = {
            'x-rapidapi-key': WORDS_API_KEY,
            'x-rapidapi-host': "wordsapiv1.p.rapidapi.com"
        }
        response = requests.get(url, headers=headers, timeout=8)
        
        if response.status_code == 200:
            data = response.json()
            if data.get("definitions") and len(data["definitions"]) > 0:
                return data["definitions"][0].get("definition", "No definition found.")
            else:
                return "No definition found for this word."
        elif response.status_code == 404:
            return "Word not found in dictionary."
        else:
            return f"API error (Status: {response.status_code})"
    except Exception as e:
        return "Definition not available at this time. (API temporarily unavailable)"

def extract_date_from_path(file_path):
    try:
        match = re.search(r'(\w+\s+\d{4})', file_path)
        if match:
            date_str = match.group(1)
            return datetime.strptime(date_str, "%b %Y")
    except:
        pass
    return datetime.min

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

if st.button("🔍 Search", type="primary"):
    if not search_word:
        st.warning("Please enter a word.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)

        if results:
            st.success(f"✅ Found {match_count} matches in {file_count} files.")

            # Dictionary Definition at the top
            definition = get_word_definition(search_word)
            st.info(f"**📖 Dictionary Definition of '{search_word}':** {definition}")

            # Sort results by folder date (newest first)
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)

            # Top Banner
            top_banner = "Newest banner.png"
            if os.path.exists(top_banner):
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.image(top_banner, width=620)

            st.subheader("📋 Search Results")
            
            for i, res in enumerate(results):
                highlighted = re.sub(
                    rf'(?<!\w){re.escape(search_word)}(?!\w)', 
                    f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>', 
                    res['text'], 
                    flags=re.IGNORECASE
                )
                
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""
                    <div style="font-family: Calibri, Arial, sans-serif; 
                                font-size: 0.92em; 
                                line-height: 1.75; 
                                background-color: #FFCCE0; 
                                padding: 18px; 
                                border-radius: 10px; 
                                border-left: 6px solid #D81B60; 
                                color: #1e1e2e;">
                        {highlighted}
                    </div>
                    """, unsafe_allow_html=True)

            # Download Full Report
            doc = Document()
            for section in doc.sections:
                section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)

            doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)

            for res in results:
                doc.add_paragraph(res["file"], style='Heading 3')
                p = doc.add_paragraph(res["text"])
                for run in p.runs:
                    run.italic = True

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download Full Report (Word Document)",
                        data=f,
                        file_name=f"Jesus speaks about {search_word}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            # Bottom Banner
            bottom_banner = "Bottom banner Std.png"
            if os.path.exists(bottom_banner):
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.image(bottom_banner, width=620)
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built for the community")
