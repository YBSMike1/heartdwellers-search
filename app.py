import os
import re
import streamlit as st
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.oxml import parse_xml
import tempfile
from datetime import datetime

st.set_page_config(page_title="Heartdwellers Search Tool", layout="wide")
st.title("❤️ Heartdwellers Search Tool")
st.markdown("**Search Jesus' messages to Mother Clare**")

# Path to your docx files on the server (you will set this)
# Path to your docx files
DOCX_FOLDER = st.text_input(
    "Full path to your Heartdwellers .docx files",
    value="Heartdwellers Docxs",   # We'll change this after uploading
    help="This must match the folder name you upload"
)

seen_italic_texts = set()

def highlight_run(run):
    try:
        highlight_xml = '<w:highlight w:val="yellow" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
        run._r.get_or_add_rPr().append(parse_xml(highlight_xml))
    except:
        pass

def search_italic_text(search_word, folder_path):
    if not os.path.exists(folder_path):
        st.error(f"Folder not found: {folder_path}")
        return [], 0, 0

    results = []
    file_count = 0
    match_count = 0
    pattern = re.compile(rf'(?<!\w){re.escape(search_word)}(?!\w)', re.IGNORECASE)

    progress_bar = st.progress(0)
    status = st.empty()

    total_files = sum(1 for _, _, files in os.walk(folder_path) if any(f.lower().endswith('.docx') for f in files))

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
            progress_bar.progress(processed / max(total_files, 1))
            status.text(f"Processing: {filename}")

            try:
                doc = Document(file_path)
                for p in doc.paragraphs:
                    italic_text = "".join(run.text for run in p.runs if run.italic)
                    if italic_text and re.search(pattern, italic_text):
                        italic_text = italic_text.strip()
                        if italic_text not in seen_italic_texts:
                            seen_italic_texts.add(italic_text)
                            results.append({
                                "file": os.path.relpath(file_path, folder_path),
                                "text": italic_text
                            })
                            match_count += 1
            except:
                continue

    progress_bar.progress(1.0)
    return results, file_count, match_count

def generate_word_document(results, search_word):
    doc = Document()
    for section in doc.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    doc.add_heading(f'What did Jesus teach us about "{search_word}"?', level=1)

    for result in results:
        doc.add_paragraph(result["file"], style='Heading 3')
        para = doc.add_paragraph()
        words = result["text"].split()
        for word in words:
            run = para.add_run(word + " ")
            run.italic = True
            if re.search(rf'(?<!\w){re.escape(search_word)}(?!\w)', word, re.IGNORECASE):
                highlight_run(run)

    # Summary (simple version)
    doc.add_page_break()
    doc.add_heading("Summary", level=2)
    summary = f"In Jesus' messages to Mother Clare, the word '{search_word}' appears {len(results)} times. "
    summary += "These passages highlight important spiritual truths for Heartdwellers."
    doc.add_paragraph(summary)

    # Save to temporary file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        doc.save(tmp.name)
        return tmp.name

# Main Search Interface
search_word = st.text_input("Enter the word or phrase to search:", placeholder="e.g. rapture, love, faith")

if st.button("Search", type="primary"):
    if not search_word:
        st.warning("Please enter a word to search.")
    elif not os.path.exists(DOCX_FOLDER):
        st.error("The folder path is not valid. Please check the path.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)

        if results:
            st.success(f"Found {match_count} matches in {file_count} files.")

            # Display results
            for res in results[:20]:  # Limit display
                with st.expander(res["file"]):
                    st.write(res["text"])

            # Download button
            docx_path = generate_word_document(results, search_word)
            with open(docx_path, "rb") as f:
                st.download_button(
                    label="📥 Download Full Report (Word Document)",
                    data=f,
                    file_name=f"Jesus speaks about {search_word}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            os.unlink(docx_path)  # Clean up
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built with love for the community")
