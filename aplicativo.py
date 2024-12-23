import streamlit as st
from pdf2docx import Converter
import os

# Função para converter o PDF para DOCX
def convert_pdf_to_docx(pdf_file):
    output_path = "output.docx"  # Salva como um arquivo temporário
    converter = Converter(pdf_file)
    converter.convert(output_path, start=0, end=None)
    converter.close()
    return output_path

# Interface do Streamlit
st.title("Conversor de PDF para DOCX")

# Upload de arquivo
pdf_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

if pdf_file:
    # Salve o arquivo temporário no disco
    with open("temp.pdf", "wb") as f:
        f.write(pdf_file.read())

    # Converte o PDF para DOCX
    output_file = convert_pdf_to_docx("temp.pdf")

    # Permite o download do arquivo DOCX convertido
    with open(output_file, "rb") as f:
        st.download_button("Baixar DOCX", f, file_name="output.docx")
    
    # Apaga o arquivo temporário após o processo
    os.remove("temp.pdf")
    os.remove(output_file)
