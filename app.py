import streamlit as st
import requests
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
from pypdf import PdfReader

st.set_page_config(page_title="PO Data Extractor", page_icon="📄")
st.title("📄 Purchase Order Data Extractor (Unrestricted HK Engine)")
st.write("Upload a PO document to extract items into a structured MingLiU table format.")

def style_text_element(paragraph, text, size_pt=9, bold=False):
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    
    run = paragraph.add_run(str(text))
    run.bold = bold
    run.font.name = 'MingLiU'
    run.font.size = Pt(size_pt)
    
    rPr = run._r.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), 'MingLiU')
    rPr.append(rFonts)

if "OPENROUTER_API_KEY" in st.secrets:
    api_key = st.secrets["OPENROUTER_API_KEY"]
else:
    api_key = None

if api_key:
    uploaded_file = st.file_uploader("Upload Purchase Order (PDF Only)", type=["pdf"])
    
    if uploaded_file is not None:
        if st.button("Process Document and Generate File"):
            with st.spinner("Extracting data via unrestricted cloud channel..."):
                try:
                    file_bytes = uploaded_file.read()
                    
                    # Read the PDF text structure natively in Python
                    pdf_file = io.BytesIO(file_bytes)
                    reader = PdfReader(pdf_file)
                    extracted_text = ""
                    
                    for page in reader.pages:
                        extracted_text += page.extract_text() + "\n"
                    
                    if not extracted_text.strip():
                        st.error("The uploaded PDF seems to be an image/scanned document with no text layer. Please use a digital PDF.")
                        st.stop()
                    
                    prompt = f"""
                    You are a data entry assistant. Analyze the raw text extracted from a purchase order below.
                    
                    On the very first line of your output text, extract the main header information in exactly this format:
                    HEADER | Restaurant Name | PO Number | Date
                    
                    After that first header line, list all items grouped by their department.
                    For each item, extract exactly these fields separated by a pipe character (|):
                    Department | Chinese Item Name | English Name & Spec | Qty | Price | Total
                    
                    Do not include markdown table structures, introduction sentences, or code blocks. Just output raw text lines separated by |.
                    
                    Raw Document Text:
                    {extracted_text}
                    """
                    
                    headers = {
                        "Authorization": f"Bearer {api_key}",
                        "Content-Type": "application/json",
                        "HTTP-Referer": "https://streamlit.io", 
                        "X-Title": "PO Extractor"
                    }
                    
                    payload = {
                        "model": "deepseek/deepseek-chat",
                        "messages": [
                            {
                                "role": "user",
                                "content": prompt
                            }
                        ]
                    }
                    
                    response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
                    
                    # SAFEGUARD: Capture and display any network errors transparently before decoding JSON
                    if response.status_code != 200:
                        st.error(f"OpenRouter Connection Error (Status {response.status_code}): {response.text}")
                        st.stop()
                        
                    try:
                        response_json = response.json()
                    except Exception:
                        st.error(f"Failed to parse server response as data. Raw Server output: {response.text}")
                        st.stop()
                    
                    if "choices" not in response_json:
                        st.error(f"API key verified, but server dropped request logic: {response_json}")
                        st.stop()
                        
                    ai_output = response_json["choices"][0]["message"]["content"]
                    
                    doc = Document()
                    col_widths = [Inches(1.0), Inches(1.3), Inches(2.2), Inches(0.6), Inches(0.7), Inches(0.7)]
                    
                    headers_list = ['Department', 'Chinese Item Name', 'English Translation & Specs', 'Qty', 'Price', 'Total']
                    table = doc.add_table(rows=1, cols=6)
                    table.style = 'Table Grid'
                    table.allow_autofit = False
                    
                    hdr_cells = table.rows.cells
                    for i, title in enumerate(headers_list):
                        hdr_cells[i].width = col_widths[i]
                        style_text_element(hdr_cells[i].paragraphs[0], title, size_pt=9, bold=True)
                    
                    lines = ai_output.strip().split('\n')
                    grand_total = 0.0
                    
                    for line in lines:
                        if '|' in line:
                            parts = [p.strip() for p in line.split('|')]
                            
                            # CRITICAL FIX: Targeted exact positional slice indices (e.g. parts[1]) instead of pushing a raw list
                            if "HEADER" in parts[0] and len(parts) >= 4:
                                p1 = doc.add_paragraph()
                                style_text_element(p1, f"Restaurant Name: {parts[1]}", size_pt=11, bold=True)
                                p2 = doc.add_paragraph()
                                style_text_element(p2, f"Purchase Order #: {parts[2]}", size_pt=11, bold=True)
                                p3 = doc.add_paragraph()
                                style_text_element(p3, f"Date: {parts[3]}", size_pt=11, bold=True)
                                doc.add_paragraph("")
                            
                            elif len(parts) == 6:
                                row_cells = table.add_row().cells
                                for i in range(6):
                                    row_cells[i].width = col_widths[i]
                                    style_text_element(row_cells[i].paragraphs[0], parts[i], size_pt=9, bold=False)
                                
                                try:
                                    clean_total_str = re.sub(r'[^\d.]', '', parts[5])
                                    if clean_total_str:
                                        grand_total += float(clean_total_str)
                                except ValueError:
                                    pass
                    
                    # Generate Totals calculation grid footer properties
                    footer_row = table.add_row()
                    footer_cells = footer_row.cells
                    for i in range(6):
                        footer_cells[i].width = col_widths[i]
                        
                    style_text_element(footer_cells[0].paragraphs[0], "Grand Total", size_pt=9, bold=True)
                    style_text_element(footer_cells[5].paragraphs[0], f"${grand_total:,.2f}", size_pt=9, bold=True)
                    
                    # Force strict line height padding to 0 properties
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                paragraph.paragraph_format.space_before = Pt(0)
                                paragraph.paragraph_format.space_after = Pt(0)
                                paragraph.paragraph_format.line_spacing = 1.0
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    st.success("Extraction Complete!")
                    
                    st.download_button(
                        label="📥 Download Extracted Word Document",
                        data=doc_buffer,
                        file_name="extracted_po_summary.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
                    st.error(f"An error occurred: {e}")
else:
    st.error("Missing OpenRouter API Key. Please add it to your Streamlit Cloud Secrets settings.")
