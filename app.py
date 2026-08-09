import io
import streamlit as st
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Precision Styler Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Universal Template Styler")
st.write("Upload your reference `.docx` file, paste your text, and automatically transfer styles!")

# ---------------------------------------------------------
# 1. PARSER ENGINE WITH FULL XML RETENTION
# ---------------------------------------------------------
class UniversalTemplateEngine:
    def __init__(self, docx_file):
        # Open the uploaded document directly to keep its theme, margins, and styles.xml intact
        self.doc = Document(docx_file)
        self.heading_p = None
        self.body_p = None
        self._find_exemplars()

    def _find_exemplars(self):
        """Scans the template for representative Heading and Body paragraphs."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            # Check for bold runs or short text lines
            is_bold = any(r.bold for r in p.runs if r.bold is not None)
            
            if (is_bold or len(text) < 45 or p.style.name.startswith("Heading")) and self.heading_p is None:
                self.heading_p = p
            elif len(text) >= 45 and self.body_p is None:
                self.body_p = p

        # Fallbacks
        if self.heading_p is None and len(self.doc.paragraphs) > 0:
            self.heading_p = self.doc.paragraphs[0]
        if self.body_p is None and len(self.doc.paragraphs) > 0:
            self.body_p = self.doc.paragraphs[-1]

    def _apply_exemplar_format(self, target_p, source_p, text, is_heading=False):
        """Applies paragraph properties, style names, and font runs directly from source paragraph."""
        if source_p is not None:
            # Copy Paragraph Style Name & Alignment
            target_p.style = source_p.style
            target_p.alignment = source_p.alignment
            
            # Copy Line Spacing & Spacing Before/After
            target_p.paragraph_format.line_spacing = source_p.paragraph_format.line_spacing
            target_p.paragraph_format.space_before = source_p.paragraph_format.space_before
            target_p.paragraph_format.space_after = source_p.paragraph_format.space_after

            # Create run and transfer character-level formatting
            run = target_p.add_run(text)
            
            if len(source_p.runs) > 0:
                ref_run = source_p.runs[0]
                if ref_run.font.name:
                    run.font.name = ref_run.font.name
                if ref_run.font.size:
                    run.font.size = ref_run.font.size
                if ref_run.font.color and ref_run.font.color.rgb:
                    run.font.color.rgb = ref_run.font.color.rgb
                
                run.bold = True if is_heading else ref_run.bold
                run.italic = ref_run.italic
            else:
                if is_heading:
                    run.bold = True
        else:
            run = target_p.add_run(text)
            if is_heading:
                run.bold = True

    def generate_formatted_report(self, raw_text):
        # 1. Clear out original text while preserving document setup & section margins
        body_elem = self.doc._body._element
        for p in list(self.doc.paragraphs):
            body_elem.remove(p._element)

        # 2. Split text into paragraphs/lines
        lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]

        for line in lines:
            # Flexible Heading Detection:
            # - Line starts with '#' or '##'
            # - Line ends with ':' or is under 60 chars with title words
            # - Line is fully capitalized
            lower_line = line.lower()
            is_heading = (
                line.startswith("#") or 
                line.endswith(":") or 
                line.isupper() or
                (len(line) < 60 and any(kw in lower_line for kw in ["title", "problem", "solution", "abstract", "introduction", "objective", "chapter"]))
            )

            clean_text = line.lstrip("#").strip()
            
            # Create new paragraph in the preserved document
            new_p = self.doc.add_paragraph()
            
            if is_heading:
                self._apply_exemplar_format(new_p, self.heading_p, clean_text, is_heading=True)
            else:
                self._apply_exemplar_format(new_p, self.body_p, clean_text, is_heading=False)

        # 3. Export to BytesIO stream
        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# 2. STREAMLIT USER INTERFACE
# ---------------------------------------------------------

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (Must be .docx)", type=["docx"])

st.subheader("2. Paste Your Content")
st.caption("Tip: Lines ending with a colon `:` or starting with `#` will be auto-formatted as Headings!")

raw_user_input = st.text_area("Paste text here:", value="", height=280)

if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        try:
            with st.spinner("Transferring exact visual styles..."):
                engine = UniversalTemplateEngine(uploaded_file)
                output_doc_stream = engine.generate_formatted_report(raw_user_input)

                st.success("Successfully formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error processing document: {str(e)}")
