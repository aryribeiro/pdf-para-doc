import PyPDF2
import streamlit as st
from docx import Document

def pdf_to_text(pdf_file):
    """Extrai o texto de um arquivo PDF."""
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text

def text_to_docx(text, docx_file):
    """Converte o texto extraído para um arquivo DOCX."""
    doc = Document()
    doc.add_paragraph(text)
    doc.save(docx_file)

def main():
    """Interface do usuário para conversão de PDF para DOCX usando Streamlit."""
    st.title("Conversor de PDF para DOCX")

    st.write("Envie seu arquivo PDF para convertê-lo em um arquivo DOCX.")

    pdf_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

    if pdf_file:
        # Extrair texto do PDF
        st.write("Extraindo texto do PDF...")
        text = pdf_to_text(pdf_file)

        if text:
            # Salvar o conteúdo extraído como um arquivo DOCX
            docx_file = "output.docx"
            text_to_docx(text, docx_file)
            st.success(f"Arquivo convertido com sucesso! Você pode baixar o arquivo DOCX abaixo.")
            st.download_button("Baixar DOCX", docx_file, file_name=docx_file)
        else:
            st.error("Erro ao extrair texto do PDF. Verifique o arquivo e tente novamente.")

if __name__ == "__main__":
    main()
