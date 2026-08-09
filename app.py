import io
import json
import streamlit as st
from docx import Document
import google.generativeai as genai

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Intelligent Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Intelligent Template Styler")
st.write("Upload a reference `.docx` template, paste any raw text or paragraphs, and let AI parse and format it automatically!")

# ---------------------------------------------------------
# 1. AI PARSER USING STABLE GEMINI-1.5-FLASH
# ---------------------------------------------------------
def ai_intelligent_parse(raw_text, api_key):
    """
    Uses Gemini LLM to parse unstructured text into semantically labeled blocks.
    Returns JSON list: [{"type": "heading1"|"heading2"|"body", "text": "..."}]
    """
    genai.configure(api_key=api_key)
    
    # Use the universally supported stable model endpoint
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    You are an expert document structure parser.
    Analyze the following raw text and divide it into structured logical blocks.
    Classify each block into one of these exact types: 'title', 'heading1', 'heading2', or 'body'.

    Return ONLY a raw JSON list of objects with the keys "type" and "text". Do not include markdown code block ticks or explanation.

    Raw Text:
    {raw_text}
    """

    response = model.generate_content(prompt)

    if not response or not response.text:
        raise RuntimeError("Failed to generate content from Gemini API.")

    clean_json = response.text.strip().lstrip("```json").rstrip("```").strip()
    return json.loads(clean_json)

# ---------------------------------------------------------
# 2. EXACT TEMPLATE STYLER ENGINE
# ---------------------------------------------------------
class IntelligentTemplateStyler:
    def __init__(self, docx_stream):
        self.doc = Document(docx_stream)
        self.heading_p = None
        self.body_p = None
        self._analyze_template()

    def _analyze_template(self):
        """Scans uploaded template for sample Heading and Body paragraphs."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            is_bold = any(run.bold for run in p.runs if run.bold is not None)
            if (is_bold or len(text) < 50) and self.heading_p is None:
                self.heading_p = p
            elif len(text) >= 50 and self.body_p is None:
                self.body_p = p

        if self.heading_p is None and len(self.doc.paragraphs) > 0:
            self.heading_p = self.doc.paragraphs[0]
        if self.body_p is None and len(self.doc.paragraphs) > 0:
            self.body_p = self.doc.paragraphs[-1]

    def _copy_font_style(self, source_run, target_run, force_bold=False):
        """Copies exact font family, size, RGB color, and weight."""
        if source_run.font.name:
            target_run.font.name = source_run.font.name
        if source_run.font.size:
            target_run.font.size = source_run.font.size
        if source_run.font.color and source_run.font.color.rgb:
            target_run.font.color.rgb = source_run.font.color.rgb
        target_run.bold = True if force_bold else source_run.bold

    def generate_formatted_doc(self, structured_blocks):
        h_style = self.heading_p.style if self.heading_p else 'Heading 1'
        b_style = self.body_p.style if self.body_p else 'Normal'

        # Clear placeholder text while keeping margins & document properties
        body_elem = self.doc._body._element
        for p in list(self.doc.paragraphs):
            body_elem.remove(p._element)

        # Inject AI-classified blocks into the document
        for block in structured_blocks:
            b_type = block.get("type", "body")
            text = block.get("text", "").strip()

            if not text:
                continue

            if b_type in ["title", "heading1", "heading2"]:
                new_p = self.doc.add_paragraph(style=h_style)
                run = new_p.add_run(text)
                if self.heading_p and len(self.heading_p.runs) > 0:
                    self._copy_font_style(self.heading_p.runs[0], run, force_bold=True)
                else:
                    run.bold = True
            else:
                new_p = self.doc.add_paragraph(style=b_style)
                run = new_p.add_run(text)
                if self.body_p and len(self.body_p.runs) > 0:
                    self._copy_font_style(self.body_p.runs[0], run)

        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# 3. STREAMLIT UI
# ---------------------------------------------------------

# Retrieve API key automatically from Streamlit Secrets
api_key = st.secrets.get("GEMINI_API_KEY", "")

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (Must be .docx)", type=["docx"])

st.subheader("2. Paste Raw Content")
raw_user_input = st.text_area("Paste any text or paragraph here:", value="", height=280)

if st.button("🚀 Intelligently Parse & Format Document"):
    if not api_key:
        st.error("⚠️ GEMINI_API_KEY secret not found! Please configure secrets in Streamlit Cloud settings.")
    elif uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        try:
            with st.spinner("🤖 Intelligent AI parsing in progress..."):
                structured_blocks = ai_intelligent_parse(raw_user_input, api_key)
                
            with st.spinner("🎨 Applying template styles..."):
                styler = IntelligentTemplateStyler(uploaded_file)
                output_doc_stream = styler.generate_formatted_doc(structured_blocks)

                st.success("Successfully parsed and formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_AI_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error during AI parsing or document formatting: {str(e)}")
