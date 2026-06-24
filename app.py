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
from spellchecker import SpellChecker
import json
import pandas as pd

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered", page_icon="❤️")

# ====================== DAY / NIGHT TOGGLE ======================
col_title, col_toggle = st.columns([5, 1])
with col_title:
    st.title("❤️ Heartdwellers Search Tool")
with col_toggle:
    dark_mode = st.toggle("🌙 Night Mode", value=True, key="theme_toggle")

st.markdown("**Search Jesus' messages to Mother Clare**")

# ====================== CSS ======================
if dark_mode:
    st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&display=swap" rel="stylesheet">
    <style>
        .stApp { background: linear-gradient(135deg, #1F1A24 0%, #2A1F35 100%); }
        .main .block-container { background: rgba(42, 37, 51, 0.92) !important; backdrop-filter: blur(22px); border-radius: 28px; padding: 2.8rem 2.2rem; }
        .fancy-header, .fancy-white {
            font-family: 'Playfair Display', serif;
            font-weight: 700;
            text-shadow: 2px 2px 5px #000000, 0 0 15px #ffffff;
            background: rgba(0,0,0,0.65);
            padding: 10px 22px;
            border-radius: 16px;
            display: block !important;
            white-space: nowrap;
            max-width: 700px;
            margin-left: auto !important;
            margin-right: auto !important;
            text-align: center;
        }
        .fancy-header { font-size: 1.18rem !important; color: #E91E63; margin-bottom: 12px; font-variation-settings: "opsz" 200; }
        .fancy-white { font-size: 1.18rem; color: #ffffff; margin-bottom: 14px; font-variation-settings: "opsz" 200; }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <link href="https://fonts.googleapis.com/css2?family=Playfair+Display:wght@700&display=swap" rel="stylesheet">
    <style>
        .stApp { background: #f8f5f0; }
        .main .block-container { background: #ffffff !important; border: 1px solid #d4c3b5; box-shadow: 0 4px 20px rgba(0,0,0,0.1); border-radius: 28px; padding: 2.8rem 2.2rem; }
        h1, h2, h3, p, label, .stMarkdown { color: #1a1a1a !important; }
        .fancy-header, .fancy-white {
            font-family: 'Playfair Display', serif;
            font-weight: 700;
            background: #2c2c2c !important;
            color: #ffffff !important;
            padding: 10px 22px;
            border-radius: 16px;
            display: block !important;
            white-space: nowrap;
            max-width: 700px;
            margin-left: auto !important;
            margin-right: auto !important;
            text-align: center;
        }
        .fancy-header { font-size: 1.18rem !important; margin-bottom: 12px; }
        .fancy-white { font-size: 1.18rem; margin-bottom: 14px; }
    </style>
    """, unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"
spell = SpellChecker()

# ====================== WORD LISTS ======================
SIN_WORDS = ["adultery", "anger", ...]  # your full list
GRACE_WORDS = ["love", "charity", ...]  # your full list

# ====================== YOUR FUNCTIONS (keep exactly as before) ======================
# Paste all your functions here: get_sin_frequencies, get_grace_frequencies, get_word_definition, extract_date_from_path, search_file, search_italic_text, build_sin_word_analysis, build_grace_word_analysis

# ====================== MAIN UI ======================
if os.path.exists("Newest banner.png"):
    st.image("Newest banner.png", use_container_width=True)

st.markdown('<div class="fancy-white">Enter a word or phrase here or select from Graces or Sins listed Below</div>', unsafe_allow_html=True)

search_word = st.text_input("Search term", placeholder="e.g. rapture, love, faith (typos ok)", label_visibility="collapsed", key="search_input")

col1, col2, col3 = st.columns([1, 1, 1])
with col2:
    search_clicked = st.button("🔍 Search", type="primary", use_container_width=True)

# Graces Table
st.markdown("---")
st.markdown('<h3 class="fancy-header">✨ Browse Graces Alphabetically (Most Used First)</h3>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; font-weight:500; margin: 8px 0 16px 0;">Click in the box next to any word in the table below to search it instantly.</p>', unsafe_allow_html=True)

if not os.path.exists("grace_word_library.json"):
    with st.spinner("Building Grace frequency cache..."):
        build_grace_word_analysis()

grace_frequencies = get_grace_frequencies()
df_grace = pd.DataFrame([{"Grace Word": word, "Frequency": grace_frequencies.get(word, 0)} for word in sorted(GRACE_WORDS)])
df_grace = df_grace.sort_values("Frequency", ascending=False)

grace_table = st.dataframe(
    df_grace,
    column_config={"Frequency": st.column_config.ProgressColumn("Frequency of usage in all messages", min_value=0, max_value=max(grace_frequencies.values()) if grace_frequencies else 500, format="%d")},
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row",
    key="grace_table"
)

# Sins Table
st.markdown("---")
st.markdown('<h3 class="fancy-header">📖 Browse Sins Alphabetically (Most Used First)</h3>', unsafe_allow_html=True)
st.markdown('<p style="text-align:center; font-weight:500; margin: 8px 0 16px 0;">Click in the box next to any word in the table below to search it instantly.</p>', unsafe_allow_html=True)

if not os.path.exists("sin_word_library.json"):
    with st.spinner("Building Sins frequency cache..."):
        build_sin_word_analysis()

sin_frequencies = get_sin_frequencies()
df_sin = pd.DataFrame([{"Sin Word": word, "Frequency": sin_frequencies.get(word, 0)} for word in sorted(SIN_WORDS)])
df_sin = df_sin.sort_values("Frequency", ascending=False)

sin_table = st.dataframe(
    df_sin,
    column_config={"Frequency": st.column_config.ProgressColumn("Frequency of usage in all messages", min_value=0, max_value=max(sin_frequencies.values()) if sin_frequencies else 438, format="%d")},
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row",
    key="sin_table"
)

# ====================== SAFE AUTO SEARCH ======================
if st.session_state.get("grace_table") and st.session_state.grace_table.get("selection", {}).get("rows"):
    row = st.session_state.grace_table["selection"]["rows"][0]
    selected_word = df_grace.iloc[row]["Grace Word"]
    st.session_state.search_input = selected_word
    if "search_trigger" not in st.session_state or st.session_state.search_trigger != selected_word:
        st.session_state.search_trigger = selected_word
        st.rerun()

if st.session_state.get("sin_table") and st.session_state.sin_table.get("selection", {}).get("rows"):
    row = st.session_state.sin_table["selection"]["rows"][0]
    selected_word = df_sin.iloc[row]["Sin Word"]
    st.session_state.search_input = selected_word
    if "search_trigger" not in st.session_state or st.session_state.search_trigger != selected_word:
        st.session_state.search_trigger = selected_word
        st.rerun()

# ====================== PERFORM SEARCH ======================
if (search_clicked and search_word) or st.session_state.get("search_trigger"):
    term = st.session_state.get("search_trigger") or search_word
    with st.spinner(f"Searching for “{term}”..."):
        results, file_count, match_count = search_italic_text(term, DOCX_FOLDER)

    if results:
        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
        definition = get_word_definition(term)
        st.info(f"**📖 Dictionary Definition of '{term}':** {definition}")

        # Your result display and download code...
        doc = Document()
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
        doc.add_heading(f'What did Jesus teach us about "{term}"?', level=1)
        for res in results:
            doc.add_paragraph(res["file"], style='Heading 3')
            p = doc.add_paragraph(res["text"])
            for run in p.runs: run.italic = True

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button(label="📥 Download Full Report (Word Document)", data=f, file_name=f"Jesus speaks about {term}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
        st.subheader("📋 Search Results")
        for res in results:
            highlighted = re.sub(rf'(?<!\w){re.escape(term)}(?!\w)', f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{term}</span>', res['text'], flags=re.IGNORECASE)
            with st.expander(f"📄 {res['file']}", expanded=True):
                st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0; font-style: italic;">{highlighted}</div>""", unsafe_allow_html=True)
    else:
        st.info("No matches found.")

if os.path.exists("Bottom banner Std.png"):
    st.image("Bottom banner Std.png", use_container_width=True)

st.markdown("""
<p style="text-align:center; color:#f8f9fa; font-size:0.95rem; margin-top:1.5rem;">
    Built by Mike F. with love for our Heartdwellers family
</p>
""", unsafe_allow_html=True)
