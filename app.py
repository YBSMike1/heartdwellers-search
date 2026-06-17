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

def get_sin_frequencies():
    freq = {}
    if os.path.exists("sin_word_library.json"):
        try:
            with open("sin_word_library.json", "r", encoding="utf-8") as f:
                data = json.load(f)
                for item in data.get("sin_words", []):
                    freq[item["Sin Word"]] = item["Frequency"]
        except:
            pass
    return freq

def get_word_definition(word):
    if not word or len(word) < 2: return "Please enter a valid word."
    try:
        url = f"https://wordsapiv1.p.rapidapi.com/words/{word.lower()}/definitions"
        headers = {'x-rapidapi-key': "e10c87331emshb838f6dd5aeb4e8p1a63dbjsn139eec4460a9", 'x-rapidapi-host': "wordsapiv1.p.rapidapi.com"}
        response = requests.get(url, headers=headers, timeout=8)
        if response.status_code == 200:
            data = response.json()
            if data.get("definitions") and len(data["definitions"]) > 0:
                return data["definitions"][0].get("definition", "No definition found.")
        return "No definition found for this word."
    except:
        return "Definition not available at this time."

def extract_date_from_path(file_path):
    try:
        match = re.search(r'(\w+\s+\d{4})', file_path)
        if match: return datetime.strptime(match.group(1), "%b %Y")
    except: pass
    return datetime.min

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

def search_italic_text(search_word, folder_path):
    results = []
    file_count = 0
    match_count = 0
    if not os.path.exists(folder_path):
        return [], 0, 0

    progress_bar = st.progress(0)
    status_text = st.empty()
    start_time = time.time()

    all_files = [os.path.join(root, f) for root, _, files in os.walk(folder_path)
                 for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
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
                eta_str = f"~{int((elapsed / files_done) * files_remaining)}s"
            else:
                eta_str = "calculating..."
            percent = int(progress * 100)
            status_text.markdown(f"**Searching** {files_done:,} / {total_files:,} files • **{percent}%** • {eta_str} remaining")

    progress_bar.progress(1.0)
    status_text.empty()
    return results, file_count, match_count

def build_sin_word_analysis():
    from collections import Counter
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
                        if word in SIN_WORDS:
                            sin_counter[word] += 1
        except:
            continue

    total = sum(sin_counter.values())
    ranked = []
    for rank, (word, freq) in enumerate(sin_counter.most_common(), 1):
        percentage = (freq / total * 100) if total > 0 else 0
        ranked.append({"Rank": rank, "Sin Word": word, "Frequency": freq, "% of Sin Mentions": round(percentage, 2)})

    sin_data = {
        "built_on": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "total_messages_scanned": len(all_files),
        "total_sin_occurrences": total,
        "sin_words": ranked
    }
    with open("sin_word_library.json", "w", encoding="utf-8") as f:
        json.dump(sin_data, f, indent=2)
    return sin_data

# ============ UI ============
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
    with col2:
        st.image("Newest banner.png", width=3400)

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

            # DOWNLOAD BUTTON AT TOP
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

            if os.path.exists("Bottom banner Std.png"):
                col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
                with col2:
                    st.image("Bottom banner Std.png", width=3400)
        else:
            corrected = spell.correction(search_word.lower())
            if corrected and corrected != search_word.lower():
                st.warning(f"No matches found for **'{search_word}'**.")
                if st.button(f"🔍 Search for “{corrected}” instead", type="primary"):
                    with st.spinner("Searching messages..."):
                        results, file_count, match_count = search_italic_text(corrected, DOCX_FOLDER)
                    if results:
                        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
            else:
                st.info("No matches found.")

# ============ DATAFRAME TABLE + SELECTBOX ============
st.markdown("---")
st.header("📖 Browse Sins Alphabetically")

st.markdown("Select a sin word below to search it instantly.")

sin_frequencies = get_sin_frequencies()
sorted_sins = sorted(SIN_WORDS)

# Build DataFrame
df_data = []
max_freq = max(sin_frequencies.values()) if sin_frequencies else 438

for sin in sorted_sins:
    freq = sin_frequencies.get(sin, 0)
    df_data.append({
        "Sin Word": sin,
        "Frequency": freq
    })

df = pd.DataFrame(df_data)

# Progress bar column (see-through style)
column_config = {
    "Frequency": st.column_config.ProgressColumn(
        "Frequency",
        help="How often this sin appears in the messages (0 to max)",
        min_value=0,
        max_value=max_freq,
        format="%d",
    )
}

st.dataframe(df, column_config=column_config, use_container_width=True, hide_index=True)

# Selectbox
selected_sin = st.selectbox(
    "Choose a sin word to search:",
    options=[""] + sorted_sins,
    index=0,
    key="sin_selectbox"
)

if selected_sin:
    with st.spinner(f"Searching for '{selected_sin}'..."):
        results, file_count, match_count = search_italic_text(selected_sin, DOCX_FOLDER)

    if results:
        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
        definition = get_word_definition(selected_sin)
        st.info(f"**📖 Dictionary Definition of '{selected_sin}':** {definition}")

        # DOWNLOAD BUTTON AT TOP
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

st.caption("Heartdwellers Search Tool — Built for the community")
