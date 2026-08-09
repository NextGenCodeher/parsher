import io
import streamlit as st
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

# Page Configuration
st.set_page_config(
    page_title="DocMind AI - XML Template Engine",
    page_icon="📄",
    layout="centered"
)

st.title("📄 DocMind AI: XML Template Transfer Engine")
st.write("Upload a reference Word document (`.docx`), paste your raw text, and transfer raw XML styles directly!")

class XMLTemplateStyler:
    def __init__(self, docx_file):
        # Load document as raw XML structure
        self.doc = Document(docx_file)
        self.heading_pPr_xml = None
        self.heading_rPr_xml = None
        self.body_pPr_xml = None
        self.body_rPr_xml = None
        
        self._extract_xml_nodes()

    def _extract_xml_nodes(self):
        """Parses the underlying OpenXML nodes directly from document paragraphs."""
        for p in self.doc.paragraphs:
            text = p.text.strip()
            if not text:
                continue

            # Check if paragraph has Paragraph Properties XML (<w:pPr>)
            pPr = p._p.get_or_add_pPr()
            
            # Find Run Properties XML (<w:rPr>) inside paragraph runs
            rPr = None
            if len(p.runs) > 0 and p.runs[0]._r.rPr is not None:
                rPr = p.runs[0]._r.rPr

            is_heading = any(run.bold for run in p.runs if run.bold) or len(text) < 50
            
            if is_heading and self.heading_pPr_xml is None:
                self.heading_pPr_xml = pPr.xml
                if rPr is not None:
                    self.heading_rPr_xml = rPr.xml
            elif not is_heading and self.body_pPr_xml is None:
                self.body_pPr_xml = pPr.xml
                if rPr is not None:
                    self.body_rPr_xml = rPr.xml

    def generate_formatted_doc(self, raw_text):
        """Creates a new document and applies raw XML properties to each paragraph."""
        new_doc = Document()

        # Transfer Section Margins (Section XML)
        for i, section in enumerate(self.doc.sections):
            if i < len(new_doc.sections):
                target_sec = new_doc.sections[i]
            else:
                target_sec = new_doc.add_section()
            
            target_sec.top_margin = section.top_margin
            target_sec.bottom_margin = section.bottom_margin
            target_sec.left_margin = section.left_margin
            target_sec.right_margin = section.right_margin

        # Parse user text lines
        lines = [line.strip() for line in raw_text.strip().split("\n") if line.strip()]

        for line in lines:
            is_heading = len(line) < 60 and (line.endswith(":") or line.startswith("#") or line.isupper())
            clean_text = line.lstrip("#").strip()

            new_p = new_doc.add_paragraph()
            
            # Select target XML properties
            pPr_xml = self.heading_pPr_xml if is_heading else self.body_pPr_xml
            rPr_xml = self.heading_rPr_xml if is_heading else self.body_rPr_xml

            # Inject Paragraph XML Properties directly
            if pPr_xml:
                new_p._p.get_or_add_pPr().append(parse_xml(pPr_xml))

            # Add text run
            new_run = new_p.add_run(clean_text)

            # Inject Run XML Properties directly (Fonts, Colors, Sizes)
            if rPr_xml:
                new_run._r.get_or_add_rPr().append(parse_xml(rPr_xml))

        # Save to memory stream
        output_stream = io.BytesIO()
        new_doc.save(output_stream)
        output_stream.seek(0)
        return output_stream

# ---------------------------------------------------------
# STREAMLIT USER INTERFACE
# ---------------------------------------------------------

st.subheader("1. Upload Reference Template (.docx)")
uploaded_file = st.file_uploader("Upload sample Word file", type=["docx"])

st.subheader("2. Paste Your Entire Document Content")
raw_user_input = st.text_area("Paste text here (headings, paragraphs, etc.):", value="", height=280)

if st.button("🚀 Format Content Using Uploaded Template"):
    if uploaded_file is None:
        st.error("Please upload a `.docx` template file first!")
    elif not raw_user_input.strip():
        st.error("Please paste or type some content into the box!")
    else:
        try:
            with st.spinner("Parsing raw OpenXML tags & applying styles..."):
                styler = XMLTemplateStyler(uploaded_file)
                output_doc_stream = styler.generate_formatted_doc(raw_user_input)

                st.success("Successfully generated XML-formatted document!")
                st.download_button(
                    label="📥 Download Formatted Word Document",
                    data=output_doc_stream,
                    file_name="DocMind_XML_Formatted_Output.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
        except Exception as e:
            st.error(f"Error parsing XML: {str(e)}")
