import io
import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor, Inches

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Visual Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Visual Template Transfer Engine")
st.write("Upload a reference Word document (`.docx`), paste your entire raw text in the box below, and generate a perfectly formatted report!")

# ---------------------------------------------------------
# 1. DYNAMIC TEMPLATE PARSER & STYLER ENGINE
# ---------------------------------------------------------
class DynamicTemplateEngine:
    def __init__(self, docx_file):
        """In-memory docx parser for Streamlit streams."""
        self.doc = Document(docx_file)
        self.style_profile = {"margins": {}, "body": {}, "heading": {}}
        self._extract_styles()

    def _extract_styles(self):
        """Extracts page geometry and font styles from uploaded file."""
        sec = self.doc.sections[0]
        self.style_profile["margins"] = {
            "top": sec.top_margin,
            "bottom": sec.bottom_margin,
            "left": sec.left_margin,
            "right": sec.right_margin
        }

        for paragraph in self.doc.paragraphs:
            for run in paragraph.runs:
                font_name = run.font.name or "Times New Roman"
                font_size = run.font.size or Pt(12)
                is_bold = run.bold or False
                color = run.font.color.rgb if run.font.color else None

                if is_bold and font_size >= Pt(13):
                    if "font_name" not in self.style_profile["heading"]:
                        self.style_profile["heading"] = {
                            "font_name": font_name,
                            "font_size": font_size,
                            "bold": True,
                            "color": color
                        }
                else:
                    if "font_name" not in self.style_profile["body"] and len(paragraph.text.strip()) > 0:
                        self.style_profile["body"] = {
                            "font_name": font_name,
                            "font_size": font_size,
                            "color": color
                        }

    def generate_report(self, text_blocks):
        """Re-applies extracted style profile to a new document stream."""
        new_doc = Document()

        # Apply Margins
        sec = new_doc.sections[0]
        sec.top_margin = self.style_profile["margins"]["top"]
        sec.bottom_margin = self.style_profile["margins"]["bottom"]
        sec.left_margin = self.style_profile["margins"]["left"]
        sec.right_margin = self.style_profile["margins"]["right"]

        # Inject Content
        for block in text_blocks:
            p = new_doc.add_paragraph()
            run = p.add_run(block["text"])

            if block["type"] == "heading":
                style = self.style_profile.get("heading", self.style_profile.get("body", {}))
                run.bold = True
            else:
                style = self.style_profile.get("body", {})

            run.font.name = style.get("font_name", "Times New Roman")
            if style.get("font_size"):
                run.font.size = style.get("font_size")
            if style.get("color"):
                run.font.color.rgb = style.get("color")

        output_stream = io.BytesIO()
        new_doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# Helper function to convert raw pasted text into headings and body blocks
def parse_raw_text(raw_text):
    lines = raw_text.strip().split("\n")
    blocks = []
    
    for line in lines:
        cleaned = line.strip()
        if not cleaned:
            continue
        
        # Treat short lines or lines ending with a colon/starting with # as headings
        if len(cleaned) < 60 and (cleaned.endswith(":") or cleaned.startswith("#") or cleaned.isupper()):
            heading_text = cleaned.lstrip("#").strip()
            blocks.append({"type": "heading", "text": heading_text})
        else:
            blocks.append({"type": "body", "text": cleaned})
            
    return blocks

# ---------------------------------------------------------
# 2. STREAMLIT USER INTERFACE
# ---------------------------------------------------------

# Section 1: File Upload
st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file", type=["docx"])

# Section 2: Single Raw Text / Paste Area (Starts completely empty)
st.subheader("2. Paste Your Entire Document Content")
raw_user_input = st.text_area("Paste text here (headings, paragraphs, etc.):", value="", height=280)

# Section 3: Format & Download Button
if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        with st.spinner("Parsing layout & formatting document..."):
            engine = DynamicTemplateEngine(uploaded_file)
            
            # Convert raw text into structured blocks
            content_blocks = parse_raw_text(raw_user_input)

            # Generate formatted file stream
            output_doc_stream = engine.generate_report(content_blocks)

            st.success("Successfully formatted document!")
            st.download_button(
                label="📥 Download Formatted Word Document",
                data=output_doc_stream,
                file_name="DocMind_Formatted_Output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
