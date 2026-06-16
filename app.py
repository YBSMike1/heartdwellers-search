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
from concurrent.futures import ThreadPoolExecutor, as_completed

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# Light grey page + MAXIMUM FORCE black text in search box
st.markdown("""
<style>
    /* Light grey page */
    .stApp, .main, .block-container, body, html { background-color: #f0f0f0 !important; }
    .main .block-container { background-color: #f8f8f8 !important; border-radius: 12px; padding: 2rem; }

    /* MAXIMUM FORCE on search input - you will see what you type */
    .stTextInput input,
    .stTextInput > div > div > input,
    .stTextInput > div > input,
    input[type="text"],
    .stTextInput textarea,
    .st-emotion-cache-1g8v9r8 input {
        color: #000000 !important;
        background-color: #ffffff !important;
        border: 3px solid #222222 !important;
        font-weight: 700 !important;
        font-size: 1.15em !important;
    }
    .stTextInput input::placeholder { color: #444444 !important; }

    /* All other text readable */
    h1, h2, h3, .stMarkdown, label, .stTextInput label,
    .stText, .stSpinner, .stProgress label, .stEmpty, .stSuccess, .stInfo, .stWarning, .stError, div[data-testid="stText"] {
        color: #1e1e2e !important;
        font-weight: 600 !important;
    }

    /* Dark mode fallback */
    @media (prefers-color-scheme: dark) {
        .stApp, .main, .block-container, body, html { background-color: #2c2c2c !important; }
        .main .block-container { background-color: #3a3a3a !important; }
        h1, h2, h3, .stMarkdown, label, .stTextInput label, .stText, .stSpinner, .stProgress label, .stEmpty, .stSuccess, .stInfo, .stWarning, .stError, div[data-testid="stText"] { color: #f0f0f0 !important; }
        .stTextInput input, .stTextInput > div > div > input, input[type="text"] { color: #ffffff !important; background-color: #444444 !important; }
    }

    /* Expander styles */
    div[data-testid="stExpander"] > div > div > div > div > button { background-color: #f8e8f0 !important; border: 3px solid #D81B60 !important; }
    div[data-testid="stExpander"] div[role="region"] { background-color: #FFCCE0 !important; border-left: 6px solid #D81B60 !important; }
</style>
""", unsafe_allow_html=True)

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

DOCX_FOLDER = "Heartdwellers Docxs"
WORDS_API_KEY = "e10c87331emshb838f6dd5aeb4e8p1a63dbjsn139eec4460a9"

def get_word_definition(word):
    if not word or len(word) < 2: return "Please enter a valid word."
    try:
        url = f"https://wordsapiv1.p.rapidapi.com/words/{word.lower()}/definitions"
        headers = {'x-rapidapi-key': WORDS_API_KEY, 'x-rapidapi-host': "wordsapiv1.p.rapidapi.com"}
        response = requests.get(url, headers=headers, timeout=8)
        if response.status_code == 200:
            data = response.json()
            if data.get("definitions") and len(data["definitions"]) > 0:
                return data["definitions"][0].get("definition", "No definition found.")
        return "No definition found for this word."
    except: return "Definition not available at this time."

def extract_date_from_path(file_path):
    try:
        match = re.search(r'(\w+\s+\d{4})', file_path)
        if match: return datetime.strptime(match.group(1), "%b %Y")
    except: pass
    return datetime.min

def search_file(file_path, search_word, pattern):
    try:
        doc = Document(file_path)
        for p in doc.paragraphs:
            italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
            if italic_text and re.search(pattern, italic_text):
                return {"file": os.path.relpath(file_path, DOCX_FOLDER), "text": italic_text.strip()}
    except: pass
    return None

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
    all_files = [os.path.join(root, f) for root, _, files in os.walk(folder_path) for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
    total_files = len(all_files)
    with ThreadPoolExecutor(max_workers=12) as executor:
        future_to_file = {executor.submit(search_file, f, search_word, pattern): f for f in all_files}
        for i, future in enumerate(as_completed(future_to_file)):
            result = future.result()
            if result:
                results.append(result)
                match_count += 1
            file_count += 1
            progress_bar.progress(min((i+1)/max(total_files,1), 1.0))
            status.text(f"Searching: {os.path.basename(future_to_file[future])}")
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
            definition = get_word_definition(search_word)
            st.info(f"**📖 Dictionary Definition of '{search_word}':** {definition}")
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
            if os.path.exists("Newest banner.png"):
                col1, col2, col3 = st.columns([1,2,1])
                with col2: st.image("Newest banner.png", width=1240)
            st.subheader("📋 Search Results")
            for res in results:
                highlighted = re.sub(rf'(?<!\w){re.escape(search_word)}(?!\w)', f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>', res['text'], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.92em; line-height: 1.75; background-color: #FFCCE0; padding: 18px; border-radius: 10px; border-left: 6px solid #D81B60; color: #1e1e2e;">{highlighted}</div>""", unsafe_allow_html=True)
            doc = Document()
            for section in doc.sections: section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
            doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)
            for res in results:
                doc.add_paragraph(res["file"], style='Heading 3')
                p = doc.add_paragraph(res["text"])
                for run in p.runs: run.italic = True
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(label="📥 Download Full Report (Word Document)", data=f, file_name=f"Jesus speaks about {search_word}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            if os.path.exists("Bottom banner Std.png"):
                col1, col2, col3 = st.columns([1,2,1])
                with col2: st.image("Bottom banner Std.png", width=1240)
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built for the community")
