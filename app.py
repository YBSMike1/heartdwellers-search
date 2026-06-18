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

st.set_page_config(page_title="Heartdwellers Search Tool", layout="centered")

# ============ DARK / LIGHT MODE TOGGLE ============
col1, col2 = st.columns([6, 1])
with col2:
    dark_mode = st.toggle("🌙", value=True, key="dark_mode_toggle", help="Toggle Dark / Light mode")

# ============ DYNAMIC THEME ============
if dark_mode:
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
        h1, h2, h3, h4, h5, h6 { color: #C4457A !important; font-weight: 700; }
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
        }
    </style>
    """, unsafe_allow_html=True)
else:
    st.markdown("""
    <style>
        .stApp { background-color: #f8f1f5; }
        .main .block-container {
            background-color: #ffffff;
            border-radius: 20px;
            padding: 2.5rem 2rem;
            box-shadow: 0 8px 30px rgba(0,0,0,0.08);
            max-width: 1100px;
        }
        h1, h2, h3, h4, h5, h6 { color: #C4457A !important; font-weight: 700; }
        .stMarkdown, .stText, p, span, label, div {
            color: #2d2a33 !important;
        }
        .stTextInput input {
            background-color: #ffffff !important;
            color: #2d2a33 !important;
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
        }
    </style>
    """, unsafe_allow_html=True)

DOCX_FOLDER = "Heartdwellers Docxs"
spell = SpellChecker()

# ============ SIN WORD LIST ============
SIN_WORDS = [
    "adultery", "anger", "arrogance", "arrogant", "backbiting", "bitter", "bitterness",
    "blasphemous", "blasphemy", "boastful", "complaining", "contention", "covetousness",
    "deceit", "deception", "deceive", "discord", "division", "doubt", "doubting", "drunk",
    "envy", "envious", "falsehood", "fear", "fearful", "fornication", "fury", "gluttony",
    "gossip", "greed", "hate", "hatred", "haughty", "hypocrisy", "hypocrite", "idolatry",
    "idol", "idols", "idle", "jealous", "jealousy", "judging", "judgment", "judgmental",
    "lazy", "laziness", "lie", "lust", "lustful", "lying", "malice", "materialism",
    "murmuring", "occult", "offended", "offense", "pride", "proud", "rage", "rebellion",
    "rebellious", "revenge", "selfish", "selfishness", "slander", "sloth", "sorcery",
    "stealing", "strife", "stubborn", "stubbornness", "thief", "unbelief", "unforgiveness",
    "unforgiving", "vengeance", "witchcraft", "worldly", "worldliness", "wrath"
]

# ============ GRACE / POSITIVE WORD LIST ============
GRACE_WORDS = [
    "love", "charity", "compassion", "mercy", "grace", "faith", "hope", "joy", "peace",
    "patience", "kindness", "goodness", "faithfulness", "gentleness", "self-control",
    "humility", "humbleness", "forgiveness", "forgive", "surrender", "trust", "obedience",
    "wisdom", "understanding", "prayer", "worship", "thanksgiving", "praise", "gratitude",
    "meekness", "longsuffering", "endurance", "perseverance", "steadfastness",
    "righteousness", "holiness", "purity", "truth", "honesty", "integrity",
    "generosity", "giving", "sharing", "hospitality", "service", "servant",
    "encouragement", "edification", "unity", "harmony", "reconciliation",
    "healing", "deliverance", "salvation", "redemption", "restoration",
    "blessing", "blessed", "anointing", "presence", "intimacy", "relationship",
    "abide", "remain", "dwell", "rest", "yield", "submit", "obey",
    "loving", "kind", "gentle", "patient", "faithful", "true", "pure", "holy",
    "humble", "forgiving", "grateful", "thankful", "peaceful", "joyful", "hopeful"
]

# (All functions remain the same - get_sin_frequencies, build_grace_word_analysis, etc.)

# ============ UI ============
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    st.image("Newest banner.png", use_container_width=True)

# ============ MAIN INSTRUCTION (Purple) ============
st.markdown("""
<div style="
    font-family: Georgia, 'Times New Roman', serif;
    font-size: 1.52rem;
    font-weight: 700;
    color: #9C27B0;
    text-shadow: 
        0 2px 4px rgba(0,0,0,0.22),
        0 4px 8px rgba(156, 39, 176, 0.18);
    letter-spacing: -0.3px;
    margin-bottom: 0.4rem;
    line-height: 1.3;
">
    Enter a word or phrase here or select from Graces or Sins listed Below
</div>
""", unsafe_allow_html=True)

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

            doc = Document()
            for section in doc.sections:
                section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
            doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)
            for res in results:
                doc.add_paragraph(res["file"], style='Heading 3')
                p = doc.add_paragraph(res["text"])
                for run in p.runs: run.italic = True

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download Full Report (Word Document)",
                        data=f,
                        file_name=f"Jesus speaks about {search_word}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

            results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
            st.subheader("📋 Search Results")

            for res in results:
                highlighted = re.sub(rf'(?<!\w){re.escape(search_word)}(?!\w)', 
                                     f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>', 
                                     res['text'], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0; font-style: italic;">{highlighted}</div>""", unsafe_allow_html=True)

# ============ GRACE TABLE (Purple Header) ============
st.markdown("---")
st.markdown("### <span style='color:#9C27B0;'>✨ Browse Graces Alphabetically (Most Used First)</span>", unsafe_allow_html=True)

st.markdown("**Click in the box next to any word in the table below to search it instantly.**")

if not os.path.exists("grace_word_library.json"):
    st.markdown(
        "⏳ **Building the Grace frequency cache for the first time.**<br>"
        "This runs automatically because new messages are added often. "
        "It only needs to happen once — future visits will load instantly.",
        unsafe_allow_html=True
    )
    with st.spinner("Scanning all messages for grace words..."):
        try:
            build_grace_word_analysis()
            st.success("✅ Grace frequency cache built successfully.")
        except Exception as e:
            st.error(f"Error building cache: {e}")

grace_frequencies = get_grace_frequencies()
sorted_graces = sorted(GRACE_WORDS)

df_data_grace = []
max_freq_grace = max(grace_frequencies.values()) if grace_frequencies else 500

for word in sorted_graces:
    freq = grace_frequencies.get(word, 0)
    df_data_grace.append({"Grace Word": word, "Frequency": freq})

df_grace = pd.DataFrame(df_data_grace)
df_grace = df_grace.sort_values("Frequency", ascending=False)

column_config_grace = {
    "Frequency": st.column_config.ProgressColumn(
        "Frequency of usage in all messages",
        help="How often this grace/virtue appears across all messages",
        min_value=0,
        max_value=max_freq_grace,
        format="%d",
    )
}

grace_event = st.dataframe(
    df_grace,
    column_config=column_config_grace,
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row"
)

if grace_event.selection.rows:
    selected_row = grace_event.selection.rows[0]
    selected_grace = df_grace.iloc[selected_row]["Grace Word"]

    with st.spinner(f"Searching for '{selected_grace}'..."):
        results, file_count, match_count = search_italic_text(selected_grace, DOCX_FOLDER)

    if results:
        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
        definition = get_word_definition(selected_grace)
        st.info(f"**📖 Dictionary Definition of '{selected_grace}':** {definition}")

        doc = Document()
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
        doc.add_heading(f'What did Jesus teach us about "{selected_grace}"?', level=1)
        for res in results:
            doc.add_paragraph(res["file"], style='Heading 3')
            p = doc.add_paragraph(res["text"])
            for run in p.runs: run.italic = True

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button(
                    label="📥 Download Full Report (Word Document)",
                    data=f,
                    file_name=f"Jesus speaks about {selected_grace}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
        st.subheader("📋 Search Results")

        for res in results:
            highlighted = re.sub(rf'(?<!\w){re.escape(selected_grace)}(?!\w)', 
                                 f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{selected_grace}</span>', 
                                 res['text'], flags=re.IGNORECASE)
            with st.expander(f"📄 {res['file']}", expanded=True):
                st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0; font-style: italic;">{highlighted}</div>""", unsafe_allow_html=True)

# ============ SIN TABLE (Purple Header) ============
st.markdown("---")
st.markdown("### <span style='color:#9C27B0;'>📖 Browse Sins Alphabetically (Most Used First)</span>", unsafe_allow_html=True)

st.markdown("**Click in the box next to any word in the table below to search it instantly.**")

sin_frequencies = get_sin_frequencies()
sorted_sins = sorted(SIN_WORDS)

df_data_sin = []
max_freq_sin = max(sin_frequencies.values()) if sin_frequencies else 438

for sin in sorted_sins:
    freq = sin_frequencies.get(sin, 0)
    df_data_sin.append({"Sin Word": sin, "Frequency": freq})

df_sin = pd.DataFrame(df_data_sin)
df_sin = df_sin.sort_values("Frequency", ascending=False)

column_config_sin = {
    "Frequency": st.column_config.ProgressColumn(
        "Frequency of usage in all messages",
        help="How often this sin appears across all messages",
        min_value=0,
        max_value=max_freq_sin,
        format="%d",
    )
}

sin_event = st.dataframe(
    df_sin,
    column_config=column_config_sin,
    use_container_width=True,
    hide_index=True,
    on_select="rerun",
    selection_mode="single-row"
)

if sin_event.selection.rows:
    selected_row = sin_event.selection.rows[0]
    selected_sin = df_sin.iloc[selected_row]["Sin Word"]

    with st.spinner(f"Searching for '{selected_sin}'..."):
        results, file_count, match_count = search_italic_text(selected_sin, DOCX_FOLDER)

    if results:
        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
        definition = get_word_definition(selected_sin)
        st.info(f"**📖 Dictionary Definition of '{selected_sin}':** {definition}")

        doc = Document()
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
        doc.add_heading(f'What did Jesus teach us about "{selected_sin}"?', level=1)
        for res in results:
            doc.add_paragraph(res["file"], style='Heading 3')
            p = doc.add_paragraph(res["text"])
            for run in p.runs: run.italic = True

        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            doc.save(tmp.name)
            with open(tmp.name, "rb") as f:
                st.download_button(
                    label="📥 Download Full Report (Word Document)",
                    data=f,
                    file_name=f"Jesus speaks about {selected_sin}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
        st.subheader("📋 Search Results")

        for res in results:
            highlighted = re.sub(rf'(?<!\w){re.escape(selected_sin)}(?!\w)', 
                                 f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{selected_sin}</span>', 
                                 res['text'], flags=re.IGNORECASE)
            with st.expander(f"📄 {res['file']}", expanded=True):
                st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0; font-style: italic;">{highlighted}</div>""", unsafe_allow_html=True)

# ============ BOTTOM BANNER ============
st.markdown("---")
if os.path.exists("Bottom banner Std.png"):
    st.image("Bottom banner Std.png", use_container_width=True)

st.caption("Heartdwellers Search Tool — Built for the community")
