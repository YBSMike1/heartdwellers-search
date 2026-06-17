import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
from collections import Counter

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered", page_icon="❤️")

st.markdown("""
<style>
    .stApp { background-color: #1F1A24; }
    .main .block-container { background-color: #2A2533; border-radius: 20px; padding: 2rem; max-width: 1100px; }
    h1 { color: #C4457A; font-weight: 700; }
    .stTextInput input { background:#fff !important; color:#000 !important; border:2px solid #C4457A !important; font-size:1.1em; padding:14px; }
</style>
""", unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"
CACHE_FILE = "italic_index.json"

def extract_italics(file_path):
    try:
        doc = Document(file_path)
        italics = [run.text.strip() for p in doc.paragraphs for run in p.runs if getattr(run, 'italic', False) and run.text.strip()]
        if italics:
            return {"file": os.path.relpath(file_path, DOCX_FOLDER), "italic_text": " | ".join(italics)}
    except: pass
    return None

def build_italic_cache():
    all_files = [os.path.join(root, f) for root, _, files in os.walk(DOCX_FOLDER) 
                 for f in files if f.lower().endswith('.docx') and not any(x in f.lower() for x in ["compilation ", "~$", "eom", "all messages"])]
    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = {executor.submit(extract_italics, f): f for f in all_files}
        index = []
        for i, future in enumerate(as_completed(futures)):
            result = future.result()
            if result: index.append(result)
    with open(CACHE_FILE, "w") as f: json.dump(index, f)
    st.success(f"✅ Cache ready! {len(index)} messages indexed")
    return index

def load_cache():
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE) as f: return json.load(f)
    return None

def search_italic_text(search_word):
    cache = load_cache() or build_italic_cache()
    pattern = re.compile(rf'(?<!\w){re.escape(search_word)}(?!\w)', re.IGNORECASE)
    results = [e for e in cache if pattern.search(e["italic_text"])]
    return results, len(cache), len(results)

def get_word_definition(word):
    try:
        url = f"https://wordsapiv1.p.rapidapi.com/words/{word.lower()}/definitions"
        headers = {'x-rapidapi-key': "e10c87331emshb838f6dd5aeb4e8p1a63dbjsn139eec4460a9", 'x-rapidapi-host': "wordsapiv1.p.rapidapi.com"}
        r = requests.get(url, headers=headers, timeout=8)
        if r.status_code == 200:
            data = r.json()
            return data.get("definitions", [{}])[0].get("definition", "No definition found.")
    except: pass
    return "Definition not available at this time."

def extract_date_from_path(file_path):
    try:
        m = re.search(r'(\w+\s+\d{4})', file_path)
        if m: return datetime.strptime(m.group(1), "%b %Y")
    except: pass
    return datetime.min

# ==================== CLEAN UI - PRE-SIN WORDS VERSION ====================
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
    with col2:
        st.image("Newest banner.png", width=3400)

if st.button("🚀 Build Fast Cache"):
    build_italic_cache()

search_word = st.text_input("Search term", value=st.session_state.get("search_word", ""), placeholder="e.g. pride, love, rapture, faith, obedience, grace", label_visibility="collapsed")

col_btn1, col_btn2 = st.columns([4, 1.2])
with col_btn2:
    if st.button("🔍 Search", type="primary", use_container_width=True):
        if search_word:
            with st.spinner("Searching..."):
                results, file_count, match_count = search_italic_text(search_word)
            st.success(f"✅ Found {match_count:,} matches in {file_count:,} messages")
            definition = get_word_definition(search_word)
            st.info(f"**📖 {search_word}:** {definition}")
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
            st.subheader("📋 Search Results (click to expand)")
            for res in results:
                highlighted = re.sub(rf'(?<!\w){re.escape(search_word)}(?!\w)', f'<span style="background:#ffeb3b;color:black;font-weight:bold;">{search_word}</span>', res["italic_text"], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f'<div style="background:#241F2E;padding:20px;border-left:6px solid #C4457A;border-radius:12px;color:#F5E6F0;">{highlighted}</div>', unsafe_allow_html=True)
            doc = Document()
            for section in doc.sections: section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
            doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)
            for res in results:
                doc.add_paragraph(res["file"], style='Heading 3')
                p = doc.add_paragraph(res["italic_text"])
                for run in p.runs: run.italic = True
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button("📥 Download Full Word Report", f, f"Jesus_on_{search_word}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            if os.path.exists("Bottom banner Std.png"):
                col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
                with col2:
                    st.image("Bottom banner Std.png", width=3400)
        else:
            st.warning("Please enter a word or phrase")

st.caption("❤️ Clean version before any sin word section was added — results display is now fully restored with expanders")
