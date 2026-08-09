import io
import streamlit as st
from docx import Document

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Direct Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Direct Template Transfer Engine")
st.write("Upload a reference Word document (`.docx`), paste your raw text, and transfer styles with precision!")

class DirectTemplateMutator:
    def __init__(self, docx_stream):
        # Open uploaded document directly to retain its full XML style sheet & theme
        self.doc = Document(docx_stream)

    def generate_formatted_doc(self, raw_text):
        # 1. Capture sample Heading and Body paragraphs before clearing text
        heading_p = None
        body_p = None

        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            is_heading = any(run.bold for run in p.runs if run.bold) or len(text) < 50
            if is_heading and heading_p is None:
                heading_p = p
            elif not is_heading and body_p is None:
                body_p = p

        # Fallbacks if specific paragraphs aren't found
        if heading_p is None and len(self.doc.paragraphs) > 0:
            heading_p = self.doc.paragraphs[0]
        if body_p is None and len(self.doc.paragraphs) > 0:
            body_p = self.doc.paragraphs[-1]

        # Extract styles directly
        heading_style = heading_p.style if heading_p else 'Heading 1'
        body_style = body_p.style if body_p else 'Normal'

        # 2. Clear out old paragraphs from document XML
        body_element = self.doc._body._element
        for p in list(self.doc.paragraphs):
            body_element.remove(p._element)

        # 3. Inject new text using the template's EXACT styles
        lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]

        for line in lines:
            is_heading = len(line) < 60 and (line.endswith(":") or line.startswith("#") or line.isupper())
            clean_text = line.lstrip("#").strip()

            if is_heading:
                new_p = self.doc.add_paragraph(clean_text, style=heading_style)
                if heading_p and len(heading_p.runs) > 0 and len(new_p.runs) > 0:
                    sample_run = heading_p.runs[0]
                    target_run = new_p.runs[0]
                    if sample_run.font.name:
                        target_run.font.name = sample_run.font.name
                    if sample_run.font.size:
                        target_run.font.size = sample_run.font.size
                    if sample_run.font.color and sample_run.font.color.rgb:
                        target_run.font.color.rgb = sample_run.font.color.rgb
                    target_run.bold = True
            else:
                new_p = self.doc.add_paragraph(clean_text, style=body_style)
                if body_p and len(body_p.runs) > 0 and len(new_p.runs) > 0:
                    sample_run = body_p.runs[0]
                    target_run = new_p.runs[0]
                    if sample_run.font.name:
                        target_run.font.name = sample_run.font.name
                    if sample_run.font.size:
                        target_run.font.size = sample_run.font.size

        # 4. Save mutated document stream
        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# STREAMLIT USER INTERFACE
# ---------------------------------------------------------

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (Must be .docx)", type=["docx"])

st.subheader("2. Paste Your Entire Document Content")
raw_user_input = st.text_area("Paste text here (headings, paragraphs, etc.):", value="", height=280)

if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        try:
            with st.spinner("Applying template styles directly..."):
                mutator = DirectTemplateMutator(uploaded_file)
                output_doc_stream = mutator.generate_formatted_doc(raw_user_input)

                st.success("Successfully formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error processing document: {str(e)}")
