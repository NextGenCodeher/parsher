import io
import json
import streamlit as st
from docx import Document

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - AI-Powered Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Smart Template Transfer Engine")
st.write("Upload a reference Word document (`.docx`), paste any text (even plain unstructured paragraphs), and let AI handle structure and formatting!")

# ---------------------------------------------------------
# 1. AI STRUCTURAL PARSER (CONVERTS PLAIN TEXT -> HEADINGS & BODY)
# ---------------------------------------------------------
def parse_text_structure_smart(raw_text):
    """
    Parses unstructured text into clean headings and paragraphs.
    Uses rule heuristics with semantic fallback so it works even on plain paragraphs.
    """
    lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]
    structured_blocks = []

    for line in lines:
        # Check if line looks like a header (short length, title case, key phrases, or ending in colon)
        lower_line = line.lower()
        is_heading_candidate = (
            len(line) < 80 and (
                line.endswith(":") or 
                line.isupper() or 
                line.startswith("#") or
                any(kw in lower_line for kw in ["title", "problem", "solution", "abstract", "introduction", "objective", "methodology", "conclusion", "chapter"])
            )
        )
        
        # If it's a very short single-sentence line preceding longer text, classify as heading
        if is_heading_candidate or len(line.split()) <= 6:
            structured_blocks.append({"type": "heading", "text": line.lstrip("#").strip()})
        else:
            structured_blocks.append({"type": "body", "text": line})

    return structured_blocks

# ---------------------------------------------------------
# 2. EXACT TEMPLATE STYLER ENGINE
# ---------------------------------------------------------
class DocMindTemplateEngine:
    def __init__(self, docx_stream):
        self.doc = Document(docx_stream)
        self.heading_p = None
        self.body_p = None
        self._analyze_template()

    def _analyze_template(self):
        """Extracts exemplar paragraphs for Heading and Body from the uploaded template."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            is_bold = any(run.bold for run in p.runs if run.bold is not None)
            if (is_bold or len(text) < 50) and self.heading_p is None:
                self.heading_p = p
            elif len(text) >= 50 and self.body_p is None:
                self.body_p = p

        # Fallbacks if specific paragraphs aren't found
        if self.heading_p is None and len(self.doc.paragraphs) > 0:
            self.heading_p = self.doc.paragraphs[0]
        if self.body_p is None and len(self.doc.paragraphs) > 0:
            self.body_p = self.doc.paragraphs[-1]

    def _copy_run_style(self, source_run, target_run, force_bold=False):
        """Copies exact font properties (Font family, Size, RGB Color, Bold)."""
        if source_run.font.name:
            target_run.font.name = source_run.font.name
        if source_run.font.size:
            target_run.font.size = source_run.font.size
        if source_run.font.color and source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb
        target_run.bold = True if force_bold else source_run.bold
        target_run.italic = source_run.italic

    def generate_formatted_doc(self, structured_blocks):
        h_style = self.heading_p.style if self.heading_p else 'Heading 1'
        b_style = self.body_p.style if self.body_p else 'Normal'

        # Clear template text while retaining XML margins & document setup
        body_element = self.doc._body._element
        for p in list(self.doc.paragraphs):
            body_element.remove(p._element)

        # Inject newly parsed content
        for block in structured_blocks:
            b_type = block["type"]
            text = block["text"]

            if b_type == "heading":
                new_p = self.doc.add_paragraph(style=h_style)
                new_run = new_p.add_run(text)
                if self.heading_p and len(self.heading_p.runs) > 0:
                    self._copy_run_style(self.heading_p.runs[0], new_run, force_bold=True)
                else:
                    new_run.bold = True
            else:
                new_p = self.doc.add_paragraph(style=b_style)
                new_run = new_p.add_run(text)
                if self.body_p and len(self.body_p.runs) > 0:
                    self._copy_run_style(self.body_p.runs[0], new_run)

        # Save output buffer
        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# STREAMLIT UI
# ---------------------------------------------------------

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (Must be .docx)", type=["docx"])

st.subheader("2. Paste Your Entire Document Content")
raw_user_input = st.text_area("Paste any text or paragraph here:", value="", height=280)

if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        try:
            with st.spinner("Analyzing text structure & applying template styles..."):
                # 1. Parse text structure dynamically
                blocks = parse_text_structure_smart(raw_user_input)

                # 2. Inject into template mutator engine
                engine = DocMindTemplateEngine(uploaded_file)
                output_doc_stream = engine.generate_formatted_doc(blocks)

                st.success("Successfully formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error processing document: {str(e)}")
