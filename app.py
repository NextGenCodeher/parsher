import io
import json
import streamlit as st
from docx import Document
import google.generativeai as genai

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - Precision Template Transfer Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: Precision Template Transfer Engine")
st.write("Upload a reference `.docx` template, paste your text, and transfer styles with 100% precision!")

# ---------------------------------------------------------
# 1. AI PARSER WITH DYNAMIC MODEL RESOLUTION
# ---------------------------------------------------------
def ai_intelligent_parse(raw_text, api_key):
    """
    Parses unstructured text into semantically labeled blocks via Gemini API.
    """
    genai.configure(api_key=api_key)
    
    candidate_models = [
        'gemini-2.0-flash',
        'gemini-1.5-flash',
        'gemini-1.5-pro'
    ]
    
    try:
        available = [m.name.replace("models/", "") for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
        flash_available = [m for m in available if 'flash' in m]
        if flash_available:
            candidate_models = flash_available + candidate_models
        elif available:
            candidate_models = available + candidate_models
    except Exception:
        pass

    prompt = f"""
    You are an expert document structure parser.
    Analyze the raw text and divide it into structured logical blocks.
    Classify each block into one of these exact types: 'title', 'heading1', 'heading2', or 'body'.

    Return ONLY a raw JSON list of objects with the keys "type" and "text". Do not include markdown fences or extra prose.

    Raw Text:
    {raw_text}
    """

    response = None
    last_error = None

    for model_name in candidate_models:
        try:
            model = genai.GenerativeModel(model_name)
            res = model.generate_content(prompt)
            if res and res.text:
                response = res
                break
        except Exception as e:
            last_error = e
            continue

    if not response or not response.text:
        raise RuntimeError(f"Could not connect to Gemini API: {str(last_error)}")

    clean_json = response.text.strip().lstrip("```json").rstrip("```").strip()
    return json.loads(clean_json)

# ---------------------------------------------------------
# 2. ACCURATE IN-PLACE TEMPLATE MUTATION ENGINE
# ---------------------------------------------------------
class PreciseInPlaceEngine:
    def __init__(self, docx_stream):
        # Open template directly so XML theme, margins, and styles remain untouched
        self.doc = Document(docx_stream)
        self.style_map = {}
        self._map_exemplar_styles()

    def _map_exemplar_styles(self):
        """Identifies standard heading/body styles present in the document."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue
            
            style_name = p.style.name
            
            if 'Heading 1' in style_name or 'Title' in style_name:
                self.style_map['heading1'] = style_name
            elif 'Heading 2' in style_name or 'Subtitle' in style_name:
                self.style_map['heading2'] = style_name
            elif 'Normal' in style_name or 'Body' in style_name:
                self.style_map['body'] = style_name

        # Fallbacks to native styles
        styles_in_doc = [s.name for s in self.doc.styles]
        if 'heading1' not in self.style_map:
            self.style_map['heading1'] = 'Heading 1' if 'Heading 1' in styles_in_doc else 'Normal'
        if 'heading2' not in self.style_map:
            self.style_map['heading2'] = 'Heading 2' if 'Heading 2' in styles_in_doc else 'Normal'
        if 'body' not in self.style_map:
            self.style_map['body'] = 'Normal'

    def generate_formatted_doc(self, structured_blocks):
        """Mutates the document in-place to preserve exact layout and typography."""
        # 1. Clear out original text elements
        body_elem = self.doc._body._element
        for p in list(self.doc.paragraphs):
            body_elem.remove(p._element)

        # 2. Re-inject structured paragraphs using mapped template styles
        for block in structured_blocks:
            b_type = block.get("type", "body").lower()
            text = block.get("text", "").strip()

            if not text:
                continue

            target_style = self.style_map.get(b_type, self.style_map['body'])
            
            # Add paragraph using native template style class
            p = self.doc.add_paragraph(text, style=target_style)
            
            # Highlight titles/headings if style didn't force bold
            if b_type in ['title', 'heading1', 'heading2'] and len(p.runs) > 0:
                p.runs[0].bold = True

        output_stream = io.BytesIO()
        self.doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# 3. STREAMLIT UI
# ---------------------------------------------------------

api_key = st.secrets.get("GEMINI_API_KEY", "")

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file (Must be .docx)", type=["docx"])

st.subheader("2. Paste Content")
raw_user_input = st.text_area("Paste text or paragraphs here:", value="", height=280)

if st.button("🚀 Intelligently Parse & Format Document"):
    if not api_key:
        st.error("⚠️ GEMINI_API_KEY secret not found! Configure secrets in Streamlit Cloud settings.")
    elif uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type content into the box!")
    else:
        try:
            with st.spinner("🤖 Classifying content structure..."):
                structured_blocks = ai_intelligent_parse(raw_user_input, api_key)
                
            with st.spinner("🎨 Mutating template in-place for exact accuracy..."):
                engine = PreciseInPlaceEngine(uploaded_file)
                output_doc_stream = engine.generate_formatted_doc(structured_blocks)

                st.success("Successfully generated accurately formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_Precision_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error during parsing or document formatting: {str(e)}")
