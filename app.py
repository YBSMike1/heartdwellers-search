import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
import tempfile
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import time
import json
from collections import Counter

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# === SOFT ELEGANT THEME ===
st.markdown("""
<style>
    .stApp { background-color: #1F1A24; }
    .main .block-container {
        background-color: #2A2533;
        border-radius: 20px;
        padding: 2.5rem 2rem;
        box-shadow: 0 10px 40px rgba(0,0,0,0.35);
        max-width: 1100px;
    }
    h1 { color: #C4457A; font-weight: 700; letter-spacing: -0.5px; }
    .stTextInput input {
        background-color: #ffffff !important;
        color: #1F1A24 !important;
        border: 2px solid #C4457A !important;
        border-radius: 14px !important;
        font-size: 1.1rem !important;
        padding: 14px 18px !important;
    }
    .stButton button[kind="primary"] {
        background-color: #C4457A !important;
        color: white !important;
        border: 2px solid #E8A0B5 !important;
        border-radius: 50px !important;
        font-weight: 700 !important;
        font-size: 1.1rem !important;
        padding: 0.75rem 2.2rem !important;
        min-height: 3.4rem !important;
    }
</style>
""", unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"

# ============ BIBLICAL SIN KEYWORDS ============
SIN_KEYWORDS = {
    "pride", "proud", "arrogance", "haughty", "boastful", "arrogant",
    "lust", "lustful", "sexual immorality", "adultery", "fornication",
    "greed", "covetous", "covetousness", "materialism",
    "envy", "envious", "jealousy", "jealous",
    "anger", "wrath", "rage", "fury",
    "gossip", "slander", "backbiting", "talebearer",
    "offense", "offended", "bitterness", "bitter", "unforgiveness", "unforgiving",
    "idolatry", "idol", "idols",
    "lying", "lie", "deceit", "deception", "falsehood",
    "stealing", "thief", "robbery",
    "gluttony", "gluttonous",
    "sloth", "lazy", "laziness", "idle",
    "fear", "fearful", "unbelief", "doubt", "doubting",
    "strife", "division", "discord", "contention",
    "witchcraft", "occult", "sorcery",
    "rebellion", "rebellious",
    "hypocrisy", "hypocrite",
    "judgmental", "judging", "judgment",
    "complaining", "murmuring",
    "selfishness", "selfish",
    "worldliness", "worldly",
    "drunkenness", "drunk",
    "hatred", "hate", "malice",
    "revenge", "vengeance",
    "deception", "deceive",
    "stubbornness", "stubborn",
    "blasphemy", "blasphemous"
}

# ============ FUNCTIONS (unchanged) ============
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

    all_files = [os.path.join(root, f) for root, _, files in os.walk(folder_path) 
                 for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
    total_files = len(all_files)

    with ThreadPoolExecutor(max_workers=12) as executor:
        future_to_file = {executor.submit(lambda f, sw=search_word: search_file(f, sw), f): f for f in all_files}
        for i, future in enumerate(as_completed(future_to_file)):
            result = future.result()
            if result:
                results.append(result)
                match_count += 1
            file_count += 1
            progress = (i + 1) / max(total_files, 1)
            progress_bar.progress(progress)
            status_text.markdown(f"**Searching** {file_count:,} / {total_files:,} files • **{int(progress*100)}%**")

    progress_bar.progress(1.0)
    status_text.empty()
    return results, file_count, match_count

def search_file(file_path, search_word):
    try:
        doc = Document(file_path)
        pattern = re.compile(rf'(?<!\w){re.escape(search_word)}(?!\w)', re.IGNORECASE)
        for p in doc.paragraphs:
            italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
            if italic_text and pattern.search(italic_text):
                return {"file": os.path.relpath(file_path, DOCX_FOLDER), "text": italic_text.strip()}
    except: pass
    return None

def build_sin_word_analysis():
    # (same as before - kept for brevity)
    sin_counter = Counter()
    all_files = [os.path.join(root, f) for root, _, files in os.walk(DOCX_FOLDER) 
                 for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
    for file_path in all_files:
        try:
            doc = Document(file_path)
            for p in doc.paragraphs:
                italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
                if italic_text:
                    words = re.findall(r'\b[a-zA-Z]+\b', italic_text.lower())
                    for word in words:
                        if word in SIN_KEYWORDS:
                            sin_counter[word] += 1
        except: pass

    total = sum(sin_counter.values())
    ranked = []
    for rank, (word, freq) in enumerate(sin_counter.most_common(), 1):
        percentage = (freq / total * 100) if total > 0 else 0
        ranked.append({"Rank": rank, "Sin Word": word, "Frequency": freq, "% of Sin Mentions": round(percentage, 2)})
    
    sin_data = {
        "built_on": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_messages_scanned": len(all_files),
        "total_unique_sin_words_found": len(sin_counter),
        "total_sin_occurrences": total,
        "sin_words": ranked
    }
    with open("sin_word_library.json", "w", encoding="utf-8") as f:
        json.dump(sin_data, f, indent=2)
    return sin_data

def load_sin_word_analysis():
    if os.path.exists("sin_word_library.json"):
        with open("sin_word_library.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return None

# ============ UI ============

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    st.image("Newest banner.png", width=3400)

st.markdown("### Enter a word or phrase")

search_word = st.text_input("Search term", value=st.session_state.get("search_word", ""), placeholder="e.g. rapture, love, faith (typos ok)", label_visibility="collapsed")

if st.button("🔍 Search", type="primary", use_container_width=True):
    if search_word:
        with st.spinner("Searching..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)
        if results:
            st.success(f"✅ Found {match_count:,} matches")
            # ... (rest of search display remains the same - omitted for brevity)

# ============ SIN WORD SECTION ============
st.markdown("---")
st.header("📖 Sin Word Frequency in Jesus’ Messages")

if st.button("🔄 Build / Refresh Sin Word Analysis"):
    with st.spinner("Building..."):
        sin_data = build_sin_word_analysis()
        st.success("Done!")

sin_data = load_sin_word_analysis()

if sin_data:
    st.write(f"**Messages scanned:** {sin_data.get('total_messages_scanned', 0)} | **Unique sin words:** {sin_data['total_unique_sin_words_found']}")

    tab1, tab2 = st.tabs(["🔥 Ranked", "🔤 Alphabetical (All Words)"])

    with tab2:
        st.markdown("**Click the colored word to search**")
        all_sorted = sorted(sin_data['sin_words'], key=lambda x: x['Sin Word'])
        cols = st.columns(3)
        for i, item in enumerate(all_sorted):
            with cols[i % 3]:
                intensity = min(255, 80 + item['Frequency'] * 9)
                color = f"rgba({intensity}, 69, 122, 0.95)"
                if st.button(f"**{item['Sin Word']}** ({item['Frequency']})", key=f"sin_{i}"):
                    st.session_state['search_word'] = item['Sin Word']
                    st.session_state['auto_search'] = True
                    st.rerun()
else:
    st.info("Click Build to generate the list.")

st.caption("Heartdwellers Search Tool — Built for the community")
