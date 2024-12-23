import PyPDF2
import streamlit as st
from docx import Document
from io import BytesIO

def pdf_to_text(pdf_file):
    """Extrai o texto de um arquivo PDF."""
    pdf_reader = PyPDF2.PdfReader(pdf_file)
    text = ""
    for page_num in range(len(pdf_reader.pages)):
        page = pdf_reader.pages[page_num]
        text += page.extract_text()
    return text

def text_to_docx(text):
    """Converte o texto extraído para um arquivo DOCX em memória."""
    doc = Document()
    doc.add_paragraph(text)
    
    # Salvar o documento em memória com BytesIO
    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)  # Voltar para o início do arquivo em memória
    return docx_buffer

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
            # Converter o texto para DOCX em memória
            docx_buffer = text_to_docx(text)
            
            st.success(f"Arquivo convertido com sucesso! Você pode baixar o arquivo DOCX abaixo.")
            
            # Botão de download com o arquivo DOCX gerado em memória
            st.download_button(
                label="Baixar DOCX",
                data=docx_buffer,
                file_name="output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Erro ao extrair texto do PDF. Verifique o arquivo e tente novamente.")

if __name__ == "__main__":
    main()
