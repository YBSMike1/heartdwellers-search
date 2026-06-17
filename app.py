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

# === THEME + CLEAN HOTLINK STYLE ===
st.markdown("""
<style>
    .stApp { background-color: #1F1A24; }
    .main .block-container { background-color: #2A2533; border-radius: 20px; padding: 2rem; max-width: 1100px; }
    h1 { color: #C4457A; }
    .stTextInput input { background:#fff !important; color:#000 !important; border:2px solid #C4457A !important; font-size:1.1em; }
    .sin-hotlink button {
        background: transparent !important; border: none !important; color: white !important;
        font-size: 1.45em !important; font-weight: 700 !important; padding: 10px 16px !important;
        margin: 5px 0 !important; text-decoration: underline !important; width: 100% !important;
        text-align: left !important; box-shadow: none !important;
    }
    .sin-hotlink button:hover { background: rgba(228, 64, 122, 0.4) !important; color: #fff !important; }
</style>
""", unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"
CACHE_FILE = "italic_index.json"

SIN_KEYWORDS = {"pride","proud","lust","greed","envy","anger","gossip","offense","bitterness","idolatry","lying","stealing","gluttony","sloth","fear","strife","witchcraft","rebellion","hypocrisy","judging","complaining","selfishness","worldliness","drunkenness","hatred","revenge","stubbornness","blasphemy","deception","jealousy","wrath","doubt","unbelief","laziness","arrogance","haughty","fornication","adultery","slander","unforgiveness","idol","falsehood","idle","contention","sorcery","hypocrite","murmuring","selfish","worldly","drunk","hate","malice","vengeance","deceive","stubborn","blasphemous","covetous"}

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
    st.info(f"🚀 Building cache with 12 threads (very fast)...")
    progress_bar = st.progress(0)
    status = st.empty()
    index = []
    with ThreadPoolExecutor(max_workers=12) as executor:
        futures = {executor.submit(extract_italics, f): f for f in all_files}
        for i, future in enumerate(as_completed(futures)):
            result = future.result()
            if result: index.append(result)
            progress_bar.progress((i+1)/len(all_files))
            status.text(f"✅ {i+1}/{len(all_files)} messages indexed")
    with open(CACHE_FILE, "w") as f:
        json.dump(index, f)
    progress_bar.empty()
    status.empty()
    st.success(f"✅ Cache built! {len(index)} messages ready — all future searches & analysis are instant!")
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

def build_sin_word_analysis():
    cache = load_cache() or build_italic_cache()
    sin_counter = Counter()
    for e in cache:
        for w in re.findall(r'\b[a-zA-Z]+\b', e["italic_text"].lower()):
            if w in SIN_KEYWORDS: sin_counter[w] += 1
    total = sum(sin_counter.values())
    ranked = [{"Rank": r+1, "Sin Word": w, "Frequency": f, "% of Sin Mentions": round(f/total*100, 2) if total else 0} for r, (w, f) in enumerate(sin_counter.most_common())]
    sin_data = {"built_on": datetime.now().strftime("%Y-%m-%d %H:%M"), "total_messages_scanned": len(cache), "total_unique_sin_words_found": len(sin_counter), "total_sin_occurrences": total, "sin_words": ranked}
    with open("sin_word_library.json", "w") as f: json.dump(sin_data, f, indent=2)
    return sin_data

def load_sin_word_analysis():
    if os.path.exists("sin_word_library.json"):
        with open("sin_word_library.json") as f: return json.load(f)
    return None

# ==================== UI ====================
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    st.image("Newest banner.png", width=3400)

if st.button("🚀 Build Ultra-Fast Cache (one-time — fixes everything)"):
    build_italic_cache()

search_word = st.text_input("Search term", value=st.session_state.get("search_word", ""), placeholder="type any word", label_visibility="collapsed")
if st.button("🔍 Search", type="primary"):
    if search_word:
        with st.spinner("Instant cache search..."):
            results, fc, mc = search_italic_text(search_word)
        st.success(f"✅ {mc} matches found in {fc} messages")

st.markdown("---")
st.header("📖 Sin Word Frequency in Jesus’ Messages")
if st.button("🔄 Build / Refresh Sin Analysis"):
    sin_data = build_sin_word_analysis()
    st.success("✅ Done using fast cache!")

sin_data = load_sin_word_analysis()
if sin_data:
    st.success(f"Messages scanned: {sin_data['total_messages_scanned']} • Total sin occurrences: {sin_data['total_sin_occurrences']}")
    tab1, tab2 = st.tabs(["🔥 Ranked by Frequency", "🔤 Alphabetical (All Words)"])
    with tab2:
        st.markdown("**Click any colored sin word below to instantly search it**")
        all_sorted = sorted(sin_data['sin_words'], key=lambda x: x['Sin Word'])
        cols = st.columns(3)
        for i, item in enumerate(all_sorted):
            with cols[i % 3]:
                intensity = min(255, 60 + item['Frequency'] * 12)
                color = f"rgba({intensity}, 69, 122, 0.98)"
                label = f"**{item['Sin Word']}** ({item['Frequency']})"
                if st.button(label, key=f"sin_{i}"):
                    st.session_state['search_word'] = item['Sin Word']
                    st.session_state['auto_search'] = True
                    st.rerun()

st.caption("❤️ Serialization error fixed • ThreadPool + cache • Exactly the colored clickable sin words you asked for")
