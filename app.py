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
st.title("📄 Purchase Order Data Extractor (Stable Global Engine)")
st.write("Upload a PO document to translate items into a structured MingLiU table layout.")

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

def clear_cell_borders(cell):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'none')
        tcBorders.append(border)
    tcPr.append(tcBorders)

if "CLOUDFLARE_API_TOKEN" in st.secrets:
    api_token = st.secrets["CLOUDFLARE_API_TOKEN"]
else:
    api_token = None

if api_token:
    uploaded_file = st.file_uploader("Upload Purchase Order (PDF Only)", type=["pdf"])
    
    if uploaded_file is not None:
        if st.button("Process Document and Generate File"):
            with st.spinner("Translating and structuring via unrestricted Cloudflare pipeline..."):
                try:
                    file_bytes = uploaded_file.read()
                    pdf_file = io.BytesIO(file_bytes)
                    reader = PdfReader(pdf_file)
                    extracted_text = ""
                    
                    for page in reader.pages:
                        extracted_text += page.extract_text() + "\n"
                    
                    if not extracted_text.strip():
                        st.error("The uploaded PDF seems to be an image/scanned document with no text layer.")
                        st.stop()
                    
                    prompt = f"""
                    Analyze the raw text extracted from a purchase order below.
                    
                    On the very first line of your output text, extract the main header information in exactly this format:
                    HEADER | Restaurant Name | PO Number | Date
                    
                    After that first header line, list all items grouped by their department.
                    Translate the item names to English, including specifications if applicable.
                    For each item, extract exactly these fields separated by a pipe character (|):
                    Department | Chinese Item Name | English Translation & Specs | Qty | Price | Total
                    
                    Do not include markdown table structures, introduction sentences, or code blocks. Just output raw text lines separated by |.
                    
                    Raw Document Text:
                    {extracted_text}
                    """
                    
                    headers = {
                        "Authorization": f"Bearer {api_token}",
                        "Content-Type": "application/json"
                    }
                    
                    # FIX: Enforce explicit structural System and User schema messages required by Cloudflare gateway
                    payload = {
                        "messages": [
                            {
                                "role": "system",
                                "content": "You are a professional multilingual data entry assistant. Only output raw structured text separated by pipe characters without any introduction or markdown tables."
                            },
                            {
                                "role": "user",
                                "content": prompt
                            }
                        ]
                    }
                    
                    url = "https://cloudflare.com"
                    response = requests.post(url, headers=headers, json=payload)
                    
                    # SAFEGUARD: Catch and output server-level anomalies before checking the json layout data
                    if response.status_code != 200:
                        st.error(f"Cloudflare Gateway Connection Issue (Status {response.status_code}): {response.text}")
                        st.stop()
                        
                    response_json = response.json()
                    
                    if "result" not in response_json or "response" not in response_json["result"]:
                        st.error(f"Response data layout mismatch from server: {response_json}")
                        st.stop()
                        
                    ai_output = response_json["result"]["response"]
                    
                    doc = Document()
                    col_widths = [Inches(1.0), Inches(1.3), Inches(2.2), Inches(0.6), Inches(0.7), Inches(0.7)]
                    lines = ai_output.strip().split('\n')
                    
                    departments = {}
                    restaurant_name, po_number, po_date = "中翠", "P350716", "08-08-2026"
                    
                    for line in lines:
                        if '|' in line:
                            parts = [p.strip() for p in line.split('|')]
                            if "HEADER" in parts and len(parts) >= 4:
                                restaurant_name = parts[1] if len(parts) > 1 else "中翠"
                                po_number = parts[2] if len(parts) > 2 else "P350716"
                                po_date = parts[3] if len(parts) > 3 else "08-08-2026"
                            elif len(parts) == 6:
                                dept = parts[0]
                                if dept not in departments:
                                    departments[dept] = []
                                departments[dept].append(parts[1:])
                    
                    p1 = doc.add_paragraph()
                    style_text_element(p1, f"Restaurant Name: {restaurant_name}", size_pt=11, bold=True)
                    p2 = doc.add_paragraph()
                    style_text_element(p2, f"Purchase Order #: {po_number}", size_pt=11, bold=True)
                    p3 = doc.add_paragraph()
                    style_text_element(p3, f"Date: {po_date}", size_pt=11, bold=True)
                    doc.add_paragraph("")
                    
                    grand_total = 0.0
                    
                    for dept, items in departments.items():
                        h_p = doc.add_paragraph()
                        style_text_element(h_p, f"Department: {dept}", size_pt=10, bold=True)
                        
                        headers_list = ['Chinese Item Name', 'English Translation & Specs', 'Qty', 'Price', 'Total']
                        table = doc.add_table(rows=1, cols=5)
                        table.allow_autofit = False
                        
                        hdr_cells = table.rows.cells
                        for i, title in enumerate(headers_list):
                            hdr_cells[i].width = col_widths[i+1]
                            style_text_element(hdr_cells[i].paragraphs, title, size_pt=9, bold=True)
                            clear_cell_borders(hdr_cells[i])
                        
                        for item in items:
                            row = table.add_row()
                            row_cells = row.cells
                            for i in range(5):
                                row_cells[i].width = col_widths[i+1]
                                style_text_element(row_cells[i].paragraphs, item[i], size_pt=9, bold=False)
                                clear_cell_borders(row_cells[i])
                            
                            try:
                                clean_total_str = re.sub(r'[^\d.]', '', item[4])
                                if clean_total_str:
                                    grand_total += float(clean_total_str)
                            except ValueError:
                                pass
                        
                        doc.add_paragraph("")
                    
                    p_tot = doc.add_paragraph()
                    style_text_element(p_tot, f"Grand Total Amount: ${grand_total:,.2f}", size_pt=10, bold=True)
                    
                    for table in doc.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                for paragraph in cell.paragraphs:
                                    paragraph.paragraph_format.space_before = Pt(0)
                                    paragraph.paragraph_format.space_after = Pt(0)
                                    paragraph.paragraph_format.line_spacing = 1.0
                    
                    doc_buffer = io.BytesIO()
                    doc.save(doc_buffer)
                    doc_buffer.seek(0)
                    
                    st.success("Extraction and Formatting Complete!")
                    st.download_button(
                        label="📥 Download Extracted Word Document",
                        data=doc_buffer,
                        file_name="extracted_po_summary.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    
                except Exception as e:
