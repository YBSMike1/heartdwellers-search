import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from rapidfuzz import fuzz
import time

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# Light targeted CSS (theme handles most colors)
st.markdown("""
<style>
    /* Search input - clean white with pink border */
    .stTextInput input {
        border: 3px solid #D81B60 !important;
        border-radius: 12px !important;
        font-weight: 700 !important;
    }

    /* Search button - pill style matching website */
    .stForm button[kind="primary"] {
        background-color: #D81B60 !important;
        color: white !important;
        border: 4px solid #FF9EC1 !important;
        border-radius: 50px !important;
        font-weight: 800 !important;
        font-size: 1.15rem !important;
        padding: 0.7rem 2.4rem !important;
        min-height: 3.4rem !important;
        box-shadow: 0 6px 16px rgba(216, 27, 96, 0.5) !important;
    }
    .stForm button[kind="primary"]:hover {
        background-color: #FF4D94 !important;
        border-color: white !important;
    }

    /* Hide the "Press Enter to submit form" hint */
    .stForm [data-testid="stMarkdownContainer"] {
        display: none !important;
    }

    /* Result boxes */
    div[data-testid="stExpander"] > div > div > div > div > button {
        background-color: #5C1A60 !important;
        border: 2px solid #D81B60 !important;
        color: white !important;
    }
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #3D0A40 !important;
        border-left: 6px solid #D81B60 !important;
    }
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

def find_fuzzy_match(text, search_word, threshold=82):
    if not text or not search_word: return False, None
    search_lower = search_word.lower().strip()
    words = re.findall(r'\b\w+\b', text)
    best_ratio = 0
    best_word = None
    for w in words:
        ratio = fuzz.ratio(w.lower(), search_lower)
        if ratio > best_ratio:
            best_ratio = ratio
            best_word = w
    if best_ratio >= threshold: return True, best_word
    return False, None

def search_file(file_path, search_word):
    try:
        doc = Document(file_path)
        for p in doc.paragraphs:
            italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
            if italic_text:
                if " " in search_word:
                    if re.search(re.escape(search_word), italic_text, re.IGNORECASE):
                        return {"file": os.path.relpath(file_path, DOCX_FOLDER), "text": italic_text.strip(), "matched_word": search_word}
                else:
                    matched, matched_word = find_fuzzy_match(italic_text, search_word)
                    if matched:
                        return {"file": os.path.relpath(file_path, DOCX_FOLDER), "text": italic_text.strip(), "matched_word": matched_word}
    except: pass
    return None

def search_italic_text(search_word, folder_path):
    results = []
    file_count = 0
    match_count = 0
    if not os.path.exists(folder_path):
        st.error(f"Folder not found: {folder_path}")
        return [], 0, 0

    progress_bar = st.progress(0)
    status_text = st.empty()
    start_time = time.time()

    all_files = [os.path.join(root, f) for root, _, files in os.walk(folder_path) for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
    total_files = len(all_files)

    with ThreadPoolExecutor(max_workers=12) as executor:
        future_to_file = {executor.submit(search_file, f, search_word): f for f in all_files}
        for i, future in enumerate(as_completed(future_to_file)):
            result = future.result()
            if result:
                results.append(result)
                match_count += 1
            file_count += 1

            progress = (i + 1) / max(total_files, 1)
            progress_bar.progress(progress)

            elapsed = time.time() - start_time
            files_done = i + 1
            files_remaining = total_files - files_done

            if files_done > 0 and files_remaining > 0:
                avg_time = elapsed / files_done
                eta_seconds = avg_time * files_remaining
                if eta_seconds < 60:
                    eta_str = f"{int(eta_seconds)}s"
                else:
                    m = int(eta_seconds // 60)
                    s = int(eta_seconds % 60)
                    eta_str = f"{m}m {s}s"
            else:
                eta_str = "calculating..."

            percent = int(progress * 100)
            status_text.markdown(f"**Searching** {files_done:,} / {total_files:,} files • **{percent}%** • ~{eta_str} remaining")

    progress_bar.progress(1.0)
    status_text.markdown(f"**Search complete** — {match_count:,} matches found in {file_count:,} files")
    time.sleep(0.6)
    status_text.empty()

    return results, file_count, match_count

with st.form("search_form", clear_on_submit=False):
    search_word = st.text_input("Enter the word or phrase to search:", placeholder="e.g. rapture, love, faith (typos ok)")
    submitted = st.form_submit_button("🔍 Search", type="primary")

if submitted:
    if not search_word:
        st.warning("Please enter a word.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)
        if results:
            st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
            definition = get_word_definition(search_word)
            st.info(f"**📖 Dictionary Definition of '{search_word}':** {definition}")
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)

            if os.path.exists("Newest banner.png"):
                col1, col2, col3 = st.columns([1,2,1])
                with col2: st.image("Newest banner.png", width=1240)

            st.subheader("📋 Search Results")
            for res in results:
                word_to_highlight = res.get("matched_word", search_word)
                highlighted = re.sub(rf'(?<!\w){re.escape(word_to_highlight)}(?!\w)', f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{word_to_highlight}</span>', res['text'], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.92em; line-height: 1.75; background-color: #5C1A60; padding: 18px; border-radius: 10px; border-left: 6px solid #D81B60; color: #f5e6f0;">{highlighted}</div>""", unsafe_allow_html=True)

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
