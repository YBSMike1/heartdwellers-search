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

# Search Input with better styling for light backgrounds
search_word = st.text_input(
    "Enter the word or phrase to search:",
    placeholder="e.g. rapture, love, faith",
    key="search_input"
)

if st.button("🔍 Search", type="primary"):
    if not search_word:
        st.warning("Please enter a word.")
    else:
        with st.spinner("Searching messages..."):
            results, file_count, match_count = search_italic_text(search_word, DOCX_FOLDER)

        if results:
            st.success(f"✅ Found {match_count} matches in {file_count} files.")

            # Top Banner
            top_banner = "Newest banner.png"
            if os.path.exists(top_banner):
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.image(top_banner, width=620)

            st.subheader("📋 Search Results")
            
            for i, res in enumerate(results):
                highlighted = re.sub(
                    rf'(?<!\w){re.escape(search_word)}(?!\w)', 
                    f'<span style="background-color: #ffeb3b; color: black; font-weight: bold;">{search_word}</span>', 
                    res['text'], 
                    flags=re.IGNORECASE
                )
                
                with st.expander(f"📄 {res['file']}", expanded=True):
                    st.markdown(f"""
                    <div style="font-family: Calibri, Arial, sans-serif; 
                                font-size: 0.92em; 
                                line-height: 1.75; 
                                background-color: #FF9EC1; 
                                padding: 18px; 
                                border-radius: 10px; 
                                border-left: 6px solid #D81B60; 
                                color: #1e1e2e;">
                        {highlighted}
                    </div>
                    """, unsafe_allow_html=True)

            # Download Full Report
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

            # Bottom Banner
            bottom_banner = "Bottom banner Std.png"
            if os.path.exists(bottom_banner):
                col1, col2, col3 = st.columns([1, 2, 1])
                with col2:
                    st.image(bottom_banner, width=620)
        else:
            st.info("No matches found.")

st.caption("Heartdwellers Search Tool — Built for the community")
