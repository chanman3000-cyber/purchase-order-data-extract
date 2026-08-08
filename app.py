import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import re
from pypdf import PdfReader

st.set_page_config(page_title="PO Data Extractor", page_icon="📄")
st.title("📄 Universal Purchase Order Data Extractor")
st.write("Upload any PO document to extract, translate, and group items into a clean, borderless layout.")

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

# Look for the hidden Google API key in cloud secrets
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = None

if api_key:
    # Initialize the official Google GenAI Client
    client = genai.Client(api_key=api_key)
    uploaded_file = st.file_uploader("Upload Purchase Order (PDF Only)", type=["pdf"])
    
    if uploaded_file is not None:
        if st.button("Process Document and Generate File"):
            with st.spinner("Analyzing, translating, and structuring document via Gemini AI..."):
                try:
                    file_bytes = uploaded_file.read()
                    
                    # Read the PDF text layer on the server
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
                    Translate the item names to English, including specifications if applicable. Even if you see brand-new items, translate them contextually.
                    For each item, extract exactly these fields separated by a pipe character (|):
                    Department | Chinese Item Name | English Translation & Specs | Qty | Price | Total
                    
                    Do not include markdown table structures, introduction sentences, or code blocks. Just output raw text lines separated by |.
                    
                    Raw Document Text:
                    {extracted_text}
                    """
                    
                    # Call official Gemini 2.5 Flash endpoint (unrestricted from the US cloud host hosting the app)
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=prompt,
                    )
                    
                    ai_output = response.text
                    
                    doc = Document()
                    col_widths = [Inches(1.0), Inches(1.3), Inches(2.2), Inches(0.6), Inches(0.7), Inches(0.7)]
                    lines = ai_output.strip().split('\n')
                    
                    departments = {}
                    restaurant_name, po_number, po_date = "中翠", "P350716", "08-08-2026"
                    
                    for line in lines:
                        if '|' in line:
                            parts = [p.strip() for p in line.split('|')]
                            if "HEADER" in parts and len(parts) >= 4:
                                restaurant_name = parts[1] if parts[1] else "中翠"
                                po_number = parts[2] if parts[2] else "P350716"
                                po_date = parts[3] if parts[3] else "08-08-2026"
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
                        
                        hdr_cells = table.rows[cells]
                        for i, title in enumerate(headers_list):
                            hdr_cells[i].width = col_widths[i+1]
                            style_text_element(hdr_cells[i].paragraphs[0], title, size_pt=9, bold=True)
                            clear_cell_borders(hdr_cells[i])
                        
                        for item in items:
                            row = table.add_row()
                            row_cells = row.cells
                            for i in range(5):
                                row_cells[i].width = col_widths[i+1]
                                style_text_element(row_cells[i].paragraphs[0], item[i], size_pt=9, bold=False)
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
                    st.error(f"An error occurred: {e}")
else:
    st.error("Missing Gemini API Key. Please add GEMINI_API_KEY to your Streamlit Cloud Secrets settings.")
