import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
import tempfile

st.set_page_config(page_title="Heartdwellers Search Tool", layout="wide")
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

DOCX_FOLDER = st.text_input(
    "Path to Heartdwellers Docxs folder",
    value="Heartdwellers Docxs",
    help="Exact folder name you uploaded"
)

def search_italic_text(search_word, folder_path):
    results = []
    file_count = 0
    match_count = 0
    pattern = re.compile(rf'(?<!\w){re.escape(search_word)}(?!\w)', re.IGNORECASE)

    if not os.path.exists(folder_path):
        st.error(f"Folder not found: {folder_path}")
        return [], 0, 0

    progress_bar = st.progress(0)
    status = st.empty()

    total_files = sum(1 for _, _, files in os.walk(folder_path) 
                     for f in files if f.lower().endswith('.docx'))

    processed = 0
    for root, dirs, files in os.walk(folder_path):
        for filename in files:
            if not filename.lower().endswith('.docx'):
                continue
            if any(skip in filename.lower() for skip in ["compilation ", "~$", "eom", "all messages"]):
                continue

            file_path = os.path.join(root, filename)
            file_count += 1
            processed += 1
            progress_bar.progress(min(processed / max(total_files, 1), 1.0))
            status.text(f"Searching: {filename}")

            try:
                doc = Document(file_path)
                for p in doc.paragraphs:
                    italic_text = "".join(run.text for run in p.runs if getattr(run, 'italic', False))
                    if italic_text and re.search(pattern, italic_text):
                        italic_text = italic_text.strip()
                        if italic_text:
                            results.append({
                                "file": os.path.relpath(file_path, folder_path),
                                "text": italic_text
                            })
                            match_count += 1
            except:
                continue

    progress_bar.progress(1.0)
    return results, file_count, match_count

# ====================== SEARCH ======================
search_word = st.text_input("Enter the word or phrase to search:", placeholder="e.g. rapture, love, faith")

if st.button("🔍 Search", type="primary"):
    if not search_word:
        st.warning("Please enter a word.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)

        if results:
            st.success(f"✅ Found {match_count} matches in {file_count} files.")

            st.subheader("📋 Search Results")

            for i, res in enumerate(results):
                with st.expander(f"📄 {res['file']}", expanded=(i < 3)):
                    st.markdown(f"""
                    <div style="background-color: #2a2a2a; padding: 15px; border-radius: 8px; line-height: 1.6;">
                        <strong>{res['text']}</strong>
                    </div>
                    """, unsafe_allow_html=True)

            # Download Full Report (same nice format as desktop)
            doc = Document()
            for section in doc.sections:
                section.top_margin = section.bottom_margin = section.left_margin = section.right_margin = Inches(0.5)

            doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)

            for res in results:
                doc.add_paragraph(res["file"], style='Heading 3')
                p = doc.add_paragraph(res["text"])
                for run in p.runs:
                    run.italic = True

            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                doc.save(tmp.name)
                with open(tmp.name, "rb") as f:
                    st.download_button(
                        label="📥 Download Full Report (Word Document)",
                        data=f,
                        file_name=f"Jesus speaks about {search_word}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built for the community")
