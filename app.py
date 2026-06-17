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
    }
    .stButton button[kind="primary"]:hover {
        background-color: #E8A0B5 !important;
        color: #1F1A24 !important;
        border-color: #C4457A !important;
    }
    div[data-testid="stExpander"] > div > div > div > div > button {
        background-color: #322C40 !important;
        border: 1px solid #C4457A !important;
        border-radius: 12px !important;
        color: #F5E6F0 !important;
    }
    div[data-testid="stExpander"] div[role="region"] {
        background-color: #241F2E !important;
        border-left: 5px solid #C4457A !important;
    }
    .stProgress > div > div > div { background-color: #C4457A !important; }
</style>
""", unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"

# ============ BIBLICAL SIN KEYWORDS ============
SIN_KEYWORDS = { ... }  # (same list as before - kept for brevity)

def get_word_definition(word):
    # (same as before)
    pass

def extract_date_from_path(file_path):
    # (same as before)
    pass

def search_file(file_path, search_word):
    # (same as before)
    pass

def search_italic_text(search_word, folder_path):
    # (same as before)
    pass

# ============ SPIRITUAL SENTIMENT ANALYSIS ============

def analyze_spiritual_sentiment(text):
    """
    Analyzes the spiritual tone of a message from Jesus.
    Returns (sentiment_label, short_explanation)
    """
    text_lower = text.lower()

    # Define keyword groups for spiritual sentiment
    convicting = ["repent", "turn away", "examine your heart", "pride", "sin", "offense", 
                  "bitterness", "unforgiveness", "selfish", "worldly", "come out from"]
    warning = ["judgment", "danger", "destruction", "wrath", "consequence", "time is short", 
               "soon", "I will", "beware", "do not"]
    encouraging = ["I love you", "you are mine", "do not fear", "I am with you", "peace", 
                   "comfort", "strength", "I will help you"]
    hopeful = ["mercy", "forgive", "restore", "new beginning", "I will heal", "hope", 
               "redemption", "second chance"]
    tender = ["beloved", "my little one", "come to Me", "rest in Me", "I hold you", 
              "intimacy", "close to My heart"]
    urgent = ["now", "today", "hurry", "do not delay", "the hour is late", "act quickly"]

    scores = {
        "Convicting": sum(1 for w in convicting if w in text_lower),
        "Warning": sum(1 for w in warning if w in text_lower),
        "Encouraging": sum(1 for w in encouraging if w in text_lower),
        "Hopeful": sum(1 for w in hopeful if w in text_lower),
        "Tender": sum(1 for w in tender if w in text_lower),
        "Urgent": sum(1 for w in urgent if w in text_lower),
    }

    # Pick the strongest sentiment
    best_sentiment = max(scores, key=scores.get)
    score = scores[best_sentiment]

    if score == 0:
        return "Neutral", "This message is more instructional or narrative in tone."

    explanations = {
        "Convicting": "This message calls the reader to self-examination and repentance.",
        "Warning": "This message contains a strong warning about spiritual danger or consequences.",
        "Encouraging": "This message offers comfort, strength, and reassurance.",
        "Hopeful": "This message speaks of mercy, restoration, and new beginnings.",
        "Tender": "This message has a very personal and loving tone from Jesus.",
        "Urgent": "This message emphasizes the need to act quickly or without delay."
    }

    return best_sentiment, explanations[best_sentiment]

# ============ SIN WORD ANALYSIS (same as before) ============
# ... (build_sin_word_analysis and load_sin_word_analysis functions remain unchanged)

# ============ UI ============

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

# (Top banner + search section remain the same)

if search_clicked or st.session_state.get("auto_search", False):
    if not search_word:
        st.warning("Please enter a word or phrase.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)
        
        if "auto_search" in st.session_state:
            del st.session_state["auto_search"]
        
        if results:
            st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
            definition = get_word_definition(search_word)
            st.info(f"**📖 Dictionary Definition of '{search_word}':** {definition}")
            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)

            st.subheader("📋 Search Results")
            for res in results:
                # === SENTIMENT ANALYSIS ===
                sentiment, explanation = analyze_spiritual_sentiment(res['text'])

                highlighted = re.sub(
                    rf'(?<!\w){re.escape(search_word)}(?!\w)',
                    f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>',
                    res['text'],
                    flags=re.IGNORECASE
                )

                with st.expander(f"📄 {res['file']}", expanded=True):
                    # Sentiment badge
                    st.markdown(f"""
                    <div style="display: inline-block; background-color: #C4457A; color: white; 
                                padding: 4px 12px; border-radius: 20px; font-size: 0.85em; margin-bottom: 10px;">
                        <strong>Sentiment:</strong> {sentiment}
                    </div>
                    """, unsafe_allow_html=True)
                    
                    st.caption(explanation)

                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; 
                                font-size: 0.95em; line-height: 1.8; background-color: #241F2E; 
                                padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; 
                                color: #F5E6F0;">{highlighted}</div>""", unsafe_allow_html=True)

            # (Download button + bottom banner remain the same)

        else:
            st.info("No matches found.")

# (Sin Word Frequency section remains the same)

st.caption("Heartdwellers Search Tool — Built for the community")
