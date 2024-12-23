import fitz  # PyMuPDF
from docx import Document

def convert_pdf_to_docx(pdf_file):
    doc = Document()
    pdf_document = fitz.open(pdf_file)

    # Para cada página do PDF, extraia o texto e adicione ao documento do Word
    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text = page.get_text("text")
        doc.add_paragraph(text)

    # Salva o documento Word
    output_path = "output.docx"
    doc.save(output_path)
    return output_path
