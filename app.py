import streamlit as st
from google import genai
from docx import Document
import io

st.set_page_config(page_title="PO Data Extractor", page_icon="📄")
st.title("📄 Purchase Order Data Extractor")
st.write("Upload a PO document to extract items into a tight Word Document table.")

# Try to get the hidden API key from the cloud settings
if "GEMINI_API_KEY" in st.secrets:
    api_key = st.secrets["GEMINI_API_KEY"]
else:
    api_key = None

if api_key:
    # Initialize client using the hidden cloud key
    client = genai.Client(api_key=api_key)
    
    # File Uploader UI Widget
    uploaded_file = st.file_uploader("Upload Purchase Order (PDF, DOCX)", type=["pdf", "docx"])
    
    if uploaded_file is not None:
        if st.button("Process Document and Generate File"):
            with st.spinner("Analyzing document with Gemini AI..."):
                try:
                    file_bytes = uploaded_file.read()
                    mime_type = "application/pdf" if uploaded_file.name.endswith('.pdf') else "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    
                    uploaded_ai_file = client.files.upload(
                        file=io.BytesIO(file_bytes),
                        config={"mime_type": mime_type}
                    )
                    
                    # UPDATED PROMPT: Instructs the AI to capture header details first
                    prompt = """
                    Analyze this purchase order. 
                    
                    On the very first line of your output text, extract the main header information in exactly this format:
                    HEADER | Restaurant Name | PO Number | Date
                    
                    After that first header line, list all items grouped by their department.
                    For each item, extract exactly these fields separated by a pipe character (|):
                    Department | Chinese Item Name | English Name & Spec | Qty | Price | Total
                    
                    Do not include markdown table structures, just raw text lines separated by |.
                    """
                    
                    response = client.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[uploaded_ai_file, prompt]
                    )
                    
                    client.files.delete(name=uploaded_ai_file.name)
                    
                    doc = Document()
                    doc.add_heading('Purchase Order Data Extraction', level=1)
                    
                    # Create the main item table structure
                    headers = ['Department', 'Chinese Item Name', 'English Translation & Specs', 'Qty', 'Price (HKD)', 'Total Amount (HKD)']
                    table = doc.add_table(rows=1, cols=6)
                    table.style = 'Table Grid'
                    
                    hdr_cells = table.rows.cells
                    for i, title in enumerate(headers):
                        hdr_cells[i].text = title
                    
                    lines = response.text.strip().split('\n')
                    for line in lines:
                        if '|' in line:
                            parts = [p.strip() for p in line.split('|')]
                            
                            # Checks if this line is the special metadata header line
                            if parts[0] == "HEADER" and len(parts) == 4:
                                # Adds the metadata text at the top of the Word document
                                doc.add_paragraph(f"Restaurant Name: {parts[1]}")
                                doc.add_paragraph(f"Purchase Order #: {parts[2]}")
                                doc.add_paragraph(f"Date: {parts[3]}")
                                doc.add_paragraph("") # Blank line divider
                            
                            # Standard 6-column line items
                            elif len(parts) == 6:
                                row_cells = table.add_row().cells
                                for i in range(6):
                                    row_cells[i].text = parts[i]
                    
                    # Tighten spacing to zero
                    for row in table.rows:
                        for cell in row.cells:
                            for paragraph in cell.paragraphs:
                                paragraph.paragraph_format.space_before = 0
                                paragraph.paragraph_format.space_after = 0
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
    st.error("Missing Gemini API Key. Please add it to your Streamlit Cloud Secrets settings.")
