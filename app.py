import io
import streamlit as st
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Visual Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Visual Template Transfer Engine")
st.write("Upload your reference `.docx` template, paste your text below, and transfer styles with 100% accuracy!")

# ---------------------------------------------------------
# 1. ADVANCED TEMPLATE CLONING ENGINE
# ---------------------------------------------------------
class ExactTemplateStyler:
    def __init__(self, template_stream):
        # Open template directly as the working document
        self.doc = Document(template_stream)
        self.heading_template_paragraph = None
        self.body_template_paragraph = None
        self._find_template_exemplars()

    def _find_template_exemplars(self):
        """Finds representative Heading and Body paragraphs from the uploaded file."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            # Identify a heading exemplar (bold or short uppercase/title line)
            is_bold = any(run.bold for run in p.runs if run.bold is not None)
            if (is_bold or len(text) < 50) and self.heading_template_paragraph is None:
                self.heading_template_paragraph = p
            elif self.body_template_paragraph is None and len(text) > 20:
                self.body_template_paragraph = p

        # Fallback if specific paragraphs aren't detected
        if self.heading_template_paragraph is None and len(self.doc.paragraphs) > 0:
            self.heading_template_paragraph = self.doc.paragraphs[0]
        if self.body_template_paragraph is None and len(self.doc.paragraphs) > 0:
            self.body_template_paragraph = self.doc.paragraphs[-1]

    def _apply_run_formatting(self, source_run, target_run):
        """Copies exact font name, size, bold/italic, and color from source to target."""
        if source_run.font.name:
            target_run.font.name = source_run.font.name
        if source_run.font.size:
            target_run.font.size = source_run.font.size
        if source_run.font.color and source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb
        target_run.bold = source_run.bold
        target_run.italic = source_run.italic

    def generate_exact_formatted_doc(self, raw_text):
        """Clones the template document structure and injects new text."""
        # Create a fresh copy of the document
        new_doc = Document()
        
        # Copy section margins directly from the template
        for i, section in enumerate(self.doc.sections):
            if i < len(new_doc.sections):
                target_sec = new_doc.sections[i]
            else:
                target_sec = new_doc.add_section()
            
            target_sec.top_margin = section.top_margin
            target_sec.bottom_margin = section.bottom_margin
            target_sec.left_margin = section.left_margin
            target_sec.right_margin = section.right_margin

        # Process input lines
        lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]

        for line in lines:
            # Determine if current line is a heading
            is_heading = len(line) < 60 and (line.endswith(":") or line.startswith("#") or line.isupper())
            
            # Select reference paragraph
            exemplar_p = self.heading_template_paragraph if (is_heading and self.heading_template_paragraph) else self.body_template_paragraph

            new_p = new_doc.add_paragraph()
            
            if exemplar_p:
                # Copy paragraph properties (alignment, spacing, style)
                new_p.style = exemplar_p.style
                new_p.alignment = exemplar_p.alignment
                new_p.paragraph_format.line_spacing = exemplar_p.paragraph_format.line_spacing
                new_p.paragraph_format.space_before = exemplar_p.paragraph_format.space_before
                new_p.paragraph_format.space_after = exemplar_p.paragraph_format.space_after

                # Add text run and apply exact character formatting
                clean_text = line.lstrip("#").strip()
                new_run = new_p.add_run(clean_text)

                if len(exemplar_p.runs) > 0:
                    self._apply_run_formatting(exemplar_p.runs[0], new_run)
                if is_heading:
                    new_run.bold = True
            else:
                new_p.add_run(line)

        # Output in-memory stream
        output_stream = io.BytesIO()
        new_doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# 2. STREAMLIT INTERFACE
# ---------------------------------------------------------

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (e.g., CivicVoice_Final_Abstract_Proposal.docx)", type=["docx"])

st.subheader("2. Paste Your Entire Document Content")
raw_user_input = st.text_area("Paste text here (headings, paragraphs, etc.):", value="", height=280)

if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        with st.spinner("Extracting exact styles & re-building document..."):
            styler = ExactTemplateStyler(uploaded_file)
            output_doc_stream = styler.generate_exact_formatted_doc(raw_user_input)

            st.success("Successfully generated formatted document!")
            st.download_button(
                label="📥 Download Formatted Word Document",
                data=output_doc_stream,
                file_name="DocMind_Exact_Formatted_Output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
