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

# === SOFT ELEGANT THEME ===
st.markdown("""
<style>
    .stApp {
        background-color: #1F1A24;
    }
    
    .main .block-container {
        background-color: #2A2533;
        border-radius: 20px;
        padding: 2.5rem 2rem;
        box-shadow: 0 10px 40px rgba(0,0,0,0.35);
        max-width: 1100px;
    }

    h1 {
        color: #C4457A;
        font-weight: 700;
        letter-spacing: -0.5px;
    }

    .stTextInput input {
        background-color: #ffffff !important;
        color: #1F1A24 !important;
        border: 2px solid #C4457A !important;
        border-radius: 14px !important;
        font-size: 1.1rem !important;
        padding: 14px 18px !important;
        font-weight: 500;
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
        box-shadow: 0 6px 20px rgba(196, 69, 122, 0.35) !important;
        transition: all 0.2s ease;
    }
    .stButton button[kind="primary"]:hover {
        background-color: #E8A0B5 !important;
        color: #1F1A24 !important;
        border-color: #C4457A !important;
        transform: translateY(-1px);
    }

    div[data-testid="stExpander"] > div > div > div > div > button {
        background-color: #322C40 !important;
        border: 1px solid #C4457A !important;
        border-radius: 12px !important;
        color: #F5E6F0 !important;
        font-weight: 600;
    }
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #241F2E !important;
        border-left: 5px solid #C4457A !important;
        border-radius: 0 12px 12px 0 !important;
    }

    .stProgress > div > div > div {
        background-color: #C4457A !important;
    }
</style>
""", unsafe_allow_html=True)

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

# === TOP BANNER (Always visible, before search) ===
if os.path.exists("Newest banner.png"):
    col1, col2, col3 = st.columns([0.5, 3, 0.5])
    with col2:
        st.image("Newest banner.png", width=2800)

st.markdown("### Enter a word or phrase")

col1, col2 = st.columns([4, 1.2])
with col1:
    search_word = st.text_input("Search term", placeholder="e.g. rapture, love, faith (typos ok)", label_visibility="collapsed")
with col2:
    search_clicked = st.button("🔍 Search", type="primary", use_container_width=True)

if search_clicked:
    if not search_word:
        st.warning("Please enter a word or phrase.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)
        if results:
            st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
            definition = get_word_definition(search_word)
            st.info(f"**📖 Dictionary Definition of '{search_word}':** {definition}")
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)

            st.subheader("📋 Search Results")
            for res in results:
                word_to_highlight = res.get("matched_word", search_word)
                highlighted = re.sub(rf'(?<!\w){re.escape(word_to_highlight)}(?!\w)', f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{word_to_highlight}</span>', res['text'], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0;">{highlighted}</div>""", unsafe_allow_html=True)

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

            # === BOTTOM BANNER (After results) ===
            if os.path.exists("Bottom banner Std.png"):
                col1, col2, col3 = st.columns([0.5, 3, 0.5])
                with col2:
                    st.image("Bottom banner Std.png", width=2800)
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built for the community")
