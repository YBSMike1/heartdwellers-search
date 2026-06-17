# ============ ALPHABETICAL SIN GRID WITH BOTTOM-FILL SHADING ============
st.markdown("---")
st.header("📖 Sin Wheel – Quick Browse (A–Z)")

st.markdown("Click any sin word. The pink fill from the bottom shows how frequently it appears.")

sin_frequencies = get_sin_frequencies()
sorted_sins = sorted(SIN_WORDS)

# Find the highest frequency for scaling (fallback to 450)
max_freq = max(sin_frequencies.values()) if sin_frequencies else 450

# Create 5 columns
cols = st.columns(5)

for i, sin in enumerate(sorted_sins):
    with cols[i % 5]:
        freq = sin_frequencies.get(sin, 1)
        
        # Calculate fill percentage (1 = very little pink, max_freq = fully pink)
        fill_percent = max(3, min(100, int((freq / max_freq) * 100)))

        # Create button with gradient fill from the bottom
        button_html = f"""
        <button 
            onclick="window.parent.postMessage({{type: 'streamlit:setComponentValue', value: '{sin}'}}, '*')"
            style="
                width: 100%;
                background: linear-gradient(to top, #C4457A {fill_percent}%, #322C40 {fill_percent}%);
                color: white;
                border: 1px solid #E8A0B5;
                border-radius: 10px;
                padding: 11px 8px;
                margin: 4px 0;
                font-weight: 600;
                font-size: 0.95rem;
                cursor: pointer;
                transition: transform 0.1s ease;
            "
            onmouseover="this.style.transform='scale(1.03)'"
            onmouseout="this.style.transform='scale(1)'"
        >
            {sin}
        </button>
        """
        st.markdown(button_html, unsafe_allow_html=True)

# Handle click from the HTML buttons
if "last_clicked_sin" not in st.session_state:
    st.session_state.last_clicked_sin = None

# This detects when an HTML button was clicked
clicked_sin = st.session_state.get("last_clicked_sin")

if clicked_sin:
    st.session_state.last_clicked_sin = None
    with st.spinner("Searching messages..."):
        results, file_count, match_count = search_italic_text(clicked_sin, DOCX_FOLDER)

    if results:
        st.success(f"✅ Found {match_count:,} matches in {file_count:,} files.")
        definition = get_word_definition(clicked_sin)
        st.info(f"**📖 Dictionary Definition of '{clicked_sin}':** {definition}")

        # DOWNLOAD BUTTON AT TOP
        doc = Document()
        for section in doc.sections:
            section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)
        doc.add_heading(f'What did Jesus teach us about "{clicked_sin}"?', level=1)
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
                    file_name=f"Jesus speaks about {clicked_sin}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        results.sort(key=lambda x: extract_date_from_path(x["file"]), reverse=True)
        st.subheader("📋 Search Results")

        for res in results:
            highlighted = re.sub(rf'(?<!\w){re.escape(clicked_sin)}(?!\w)', 
                                 f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{clicked_sin}</span>', 
                                 res['text'], flags=re.IGNORECASE)
            with st.expander(f"📄 {res['file']}", expanded=True):
                st.markdown(f"""<div style="font-family: Calibri, Arial, sans-serif; font-size: 0.95em; line-height: 1.8; background-color: #241F2E; padding: 20px; border-radius: 12px; border-left: 6px solid #C4457A; color: #F5E6F0; font-style: italic;">{highlighted}</div>""", unsafe_allow_html=True)
