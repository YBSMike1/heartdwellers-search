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
SIN_KEYWORDS = { ... }  # (same list as before - kept for brevity)

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
    # (same as before)
    pass

def extract_date_from_path(file_path):
    try:
        match = re.search(r'(\w+\s+\d{4})', file_path)
        if match: return datetime.strptime(match.group(1), "%b %Y")
    except: pass
    return datetime.min
    # (same as before)
    pass

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
    # (same as before)
    pass

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

# ============ SIN WORD ANALYSIS ============

def build_sin_word_analysis():
    sin_counter = Counter()
    processed = 0

    progress_bar = st.progress(0)
    status = st.empty()

    all_files = [os.path.join(root, f) for root, _, files in os.walk(DOCX_FOLDER) 
                 for f in files if f.lower().endswith('.docx') and not any(s in f.lower() for s in ["compilation ", "~$", "eom", "all messages"])]
    total_files = len(all_files)

    for file_path in all_files:
        processed += 1
        progress_bar.progress(processed / total_files)
        status.text(f"Scanning: {os.path.basename(file_path)} ({processed}/{total_files})")

        try:
            doc = Document(file_path)
            for p in doc.paragraphs:
                italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
                if italic_text:
                    words = re.findall(r'\b[a-zA-Z]+\b', italic_text.lower())
                    for word in words:
                        if word in SIN_KEYWORDS:
                            sin_counter[word] += 1
        except:
            continue

    progress_bar.empty()
    status.empty()
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

    total_sin_occurrences = sum(sin_counter.values())
    # Pick the strongest sentiment
    best_sentiment = max(scores, key=scores.get)
    score = scores[best_sentiment]

    ranked_sins = []
    for rank, (word, freq) in enumerate(sin_counter.most_common(), 1):
        percentage = (freq / total_sin_occurrences * 100) if total_sin_occurrences > 0 else 0
        ranked_sins.append({
            "Rank": rank,
            "Sin Word": word,
            "Frequency": freq,
            "% of Sin Mentions": round(percentage, 2)
        })
    if score == 0:
        return "Neutral", "This message is more instructional or narrative in tone."

    sin_data = {
        "built_on": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "total_messages_scanned": total_files,
        "total_unique_sin_words_found": len(sin_counter),
        "total_sin_occurrences": total_sin_occurrences,
        "sin_words": ranked_sins
    explanations = {
        "Convicting": "This message calls the reader to self-examination and repentance.",
        "Warning": "This message contains a strong warning about spiritual danger or consequences.",
        "Encouraging": "This message offers comfort, strength, and reassurance.",
        "Hopeful": "This message speaks of mercy, restoration, and new beginnings.",
        "Tender": "This message has a very personal and loving tone from Jesus.",
        "Urgent": "This message emphasizes the need to act quickly or without delay."
}

    with open("sin_word_library.json", "w", encoding="utf-8") as f:
        json.dump(sin_data, f, indent=2)

    return sin_data
    return best_sentiment, explanations[best_sentiment]

def load_sin_word_analysis():
    if os.path.exists("sin_word_library.json"):
        with open("sin_word_library.json", "r", encoding="utf-8") as f:
            return json.load(f)
    return None
# ============ SIN WORD ANALYSIS (same as before) ============
# ... (build_sin_word_analysis and load_sin_word_analysis functions remain unchanged)

# ============ UI ============

st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

if os.path.exists("Newest banner.png"):
    col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
    with col2:
        st.image("Newest banner.png", width=3400)

st.markdown("### Enter a word or phrase")

search_word = st.text_input(
    "Search term",
    value=st.session_state.get("search_word", ""),
    placeholder="e.g. rapture, love, faith (typos ok)",
    label_visibility="collapsed"
)

col1, col2 = st.columns([4, 1.2])
with col2:
    search_clicked = st.button("🔍 Search", type="primary", use_container_width=True)
# (Top banner + search section remain the same)

if search_clicked or st.session_state.get("auto_search", False):
if not search_word:
@@ -274,69 +162,37 @@ def load_sin_word_analysis():

st.subheader("📋 Search Results")
for res in results:
                highlighted = re.sub(rf'(?<!\w){re.escape(search_word)}(?!\w)', f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>', res['text'], flags=re.IGNORECASE)
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0;">{highlighted}</div>""", unsafe_allow_html=True)
                # === SENTIMENT ANALYSIS ===
                sentiment, explanation = analyze_spiritual_sentiment(res['text'])

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
                    st.download_button(label="📥 Download Full Report (Word Document)", data=f, file_name=f"Jesus speaks about {search_word}.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
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

            if os.path.exists("Bottom banner Std.png"):
                col1, col2, col3 = st.columns([0.15, 3.7, 0.15])
                with col2:
                    st.image("Bottom banner Std.png", width=3400)
else:
st.info("No matches found.")

# ============ SIN WORD FREQUENCY SECTION ============
st.markdown("---")
st.header("📖 Sin Word Frequency in Jesus’ Messages")

st.markdown("""
This section shows how often **biblical sin-related words** appear in Jesus’ direct words (italic text).  
**Click any numbered word below** to automatically search for it.
""")

if st.button("🔄 Build / Refresh Sin Word Analysis", type="secondary"):
    with st.spinner("Scanning all messages for sin-related words..."):
        sin_data = build_sin_word_analysis()
        st.success(f"Sin word analysis updated on {sin_data['built_on']}")

sin_data = load_sin_word_analysis()

if sin_data:
    st.success(f"**Last updated:** {sin_data['built_on']}")
    st.write(f"**Total messages searched:** {sin_data.get('total_messages_scanned', 0):,}")
    st.write(f"**Unique sin-related words found:** {sin_data['total_unique_sin_words_found']}")
    st.write(f"**Total occurrences of sin-related words:** {sin_data['total_sin_occurrences']:,}")

    # === NUMBERED CLICKABLE SIN WORDS ===
    if sin_data['sin_words']:
        st.markdown("**Click a word to search:**")
        cols = st.columns(6)
        for i, item in enumerate(sin_data['sin_words'][:30]):
            with cols[i % 6]:
                button_label = f"{item['Rank']}. {item['Sin Word']}"
                if st.button(button_label, key=f"sin_{i}"):
                    st.session_state['search_word'] = item['Sin Word']
                    st.session_state['auto_search'] = True
                    st.rerun()

    # Full table
    import pandas as pd
    df = pd.DataFrame(sin_data['sin_words'])
    st.dataframe(df, use_container_width=True, hide_index=True)
else:
    st.info("Click the button above to build the sin word analysis from all messages.")
# (Sin Word Frequency section remains the same)

st.caption("Heartdwellers Search Tool — Built for the community")
