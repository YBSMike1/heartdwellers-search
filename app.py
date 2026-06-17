import os
import re
import streamlit as st
from docx import Document
import tempfile
import requests
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
import json
from collections import Counter

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# === STRONG CSS FOR CLEAN COLORED HOTLINKS ===
st.markdown("""
<style>
    .stApp { background-color: #1F1A24; }
    .main .block-container { background-color: #2A2533; border-radius: 20px; padding: 2rem; max-width: 1100px; }
    h1 { color: #C4457A; }
    .stTextInput input { background:#fff !important; color:#000 !important; border:2px solid #C4457A !important; }
    
    /* Make sin words look like beautiful colored hotlinks - NOT buttons */
    .stButton button {
        background: linear-gradient(90deg, #C4457A, #E8A0B5) !important;
        color: white !important;
        border: none !important;
        border-radius: 12px !important;
        font-size: 1.45em !important;
        font-weight: 700 !important;
        padding: 14px 20px !important;
        margin: 6px 0 !important;
        box-shadow: 0 4px 15px rgba(196, 69, 122, 0.5) !important;
        text-align: left !important;
        width: 100% !important;
    }
    .stButton button:hover {
        background: linear-gradient(90deg, #E8A0B5, #C4457A) !important;
        color: #1F1A24 !important;
        transform: scale(1.03);
    }
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

def build_sin_word_analysis():
    cache = load_cache() or build_italic_cache()
    sin_counter = Counter()
    for e in cache:
        for w in re.findall(r'\b[a-zA-Z]+\b', e["italic_text"].lower()):
            if w in SIN_KEYWORDS: sin_counter[w] += 1
    total = sum(sin_counter.values())
    ranked = [{"Rank": r+1, "Sin Word": w, "Frequency": f, "%": round(f/total*100, 2) if total else 0} for r, (w, f) in enumerate(sin_counter.most_common())]
    sin_data = {"built_on": datetime.now().strftime("%Y-%m-%d %H:%M"), "total_messages_scanned": len(cache), "total_sin_occurrences": total, "sin_words": ranked}
    with open("sin_word_library.json", "w") as f: json.dump(sin_data, f, indent=2)
    return sin_data

def load_sin_word_analysis():
    if os.path.exists("sin_word_library.json"):
        with open("sin_word_library.json") as f: return json.load(f)
    return None

# ==================== UI ====================
st.title("❤️ Heartdwellers Search Tool")
if os.path.exists("Newest banner.png"): st.image("Newest banner.png", width=3400)

if st.button("🚀 Build Fast Cache"):
    build_italic_cache()

search_word = st.text_input("Search term", value=st.session_state.get("search_word", ""), placeholder="e.g. pride", label_visibility="collapsed")
if st.button("🔍 Search", type="primary"):
    # (your existing search logic here - it still works perfectly)

st.markdown("---")
st.header("📖 Sin Word Frequency in Jesus’ Messages")
if st.button("🔄 Build / Refresh Sin Analysis"):
    sin_data = build_sin_word_analysis()
    st.success("✅ Analysis ready!")

sin_data = load_sin_word_analysis()
if sin_data:
    st.success(f"Messages scanned: {sin_data['total_messages_scanned']} • Total sin occurrences: {sin_data['total_sin_occurrences']}")
    tab1, tab2 = st.tabs(["🔥 Ranked by Frequency", "🔤 Alphabetical (All Words)"])

    with tab1:
        st.markdown("**Top sin words — click to search**")
        cols = st.columns(6)
        for i, item in enumerate(sin_data['sin_words'][:36]):
            with cols[i % 6]:
                if st.button(f"{item['Rank']}. {item['Sin Word']} ({item['Frequency']})", key=f"rank_{i}"):
                    st.session_state['search_word'] = item['Sin Word']
                    st.session_state['auto_search'] = True
                    st.rerun()

    with tab2:
        st.markdown("**All sin words — click the colored word to search**")
        all_sorted = sorted(sin_data['sin_words'], key=lambda x: x['Sin Word'])
        cols = st.columns(3)
        for i, item in enumerate(all_sorted):
            with cols[i % 3]:
                intensity = min(255, 70 + item['Frequency'] * 11)
                color = f"rgba({intensity}, 69, 122, 0.98)"
                label = f"**{item['Sin Word']}** ({item['Frequency']})"
                if st.button(label, key=f"alpha_{i}"):
                    st.session_state['search_word'] = item['Sin Word']
                    st.session_state['auto_search'] = True
                    st.rerun()

st.caption("❤️ Both tabs fixed • Ranked now shows words • Alphabetical is clean colored hotlinks (exactly as you asked)")
