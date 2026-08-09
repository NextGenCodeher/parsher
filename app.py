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
st.write("Upload any reference Word document (`.docx`), paste your new content, and generate a perfectly styled report!")

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
        # 1. Extract Margins
        sec = self.doc.sections[0]
        self.style_profile["margins"] = {
            "top": sec.top_margin,
            "bottom": sec.bottom_margin,
            "left": sec.left_margin,
            "right": sec.right_margin
        }

        # 2. Extract Typography & Hierarchy
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

        # Inject Text Blocks with Applied Rules
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

        # Save into an in-memory buffer for instant browser download
        output_stream = io.BytesIO()
        new_doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# 2. STREAMLIT USER INTERFACE
# ---------------------------------------------------------

# Section 1: Template Upload
st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader(
    "Upload sample Word file (e.g., CivicVoice_Final_Abstract_Proposal.docx)", 
    type=["docx"]
)

# Section 2: Text Inputs
st.subheader("2. Enter Your New Project Content")
project_title = st.text_input("Project Title:", "Project Title: DocMind AI")

problem_stmt = st.text_area(
    "Problem Statement:", 
    "Writing technical documents like college project reports, research papers, or software guides takes way too much time and effort. People often spend almost half their working time just fixing page layouts, margins, font styles, and citation numbers instead of focusing on their actual work."
)

proposed_sol = st.text_area(
    "Proposed Solution:", 
    "DocMind AI is a smart tool that handles both technical writing and document formatting automatically. It reads your real project files—like code, READMEs, and dataset notes—so every sentence it generates is completely accurate with no made-up facts. A team of three AI agents works together to build a clear table of contents and write out the chapters step by step."
)

# Section 3: Action Button
if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    else:
        with st.spinner("Extracting layout geometry & applying styles..."):
            # Initialize Engine
            engine = DynamicTemplateEngine(uploaded_file)
            
            st.success("Successfully parsed template layout!")
            col1, col2 = st.columns(2)
            col1.metric("Heading Style", f"{engine.style_profile['heading'].get('font_name', 'Default')} ({engine.style_profile['heading'].get('font_size', Pt(14)).pt}pt)")
            col2.metric("Body Style", f"{engine.style_profile['body'].get('font_name', 'Default')} ({engine.style_profile['body'].get('font_size', Pt(12)).pt}pt)")

            # Structure content payload
            content_blocks = [
                {"type": "heading", "text": project_title},
                {"type": "heading", "text": "Problem Statement:"},
                {"type": "body", "text": problem_stmt},
                {"type": "heading", "text": "Proposed Solution:"},
                {"type": "body", "text": proposed_sol}
            ]

            # Generate formatted stream
            output_doc_stream = engine.generate_report(content_blocks)

            # Download Trigger
            st.download_button(
                label="📥 Download Formatted Word Document",
                data=output_doc_stream,
                file_name="DocMind_Formatted_Output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )