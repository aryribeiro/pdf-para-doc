import streamlit as st
from pdf2docx import Converter
import os

# Função para converter PDF para DOCX
def convert_pdf_to_docx(pdf_file):
    # Verificando se o arquivo foi carregado corretamente
    if not os.path.exists(pdf_file):
        raise FileNotFoundError(f"O arquivo PDF não foi encontrado: {pdf_file}")
    
    output_path = "output.docx"  # ou qualquer caminho válido
    try:
        # Usando a biblioteca para fazer a conversão
        cv = Converter(pdf_file)
        cv.convert(output_path, start=0, end=None)  # converte o arquivo
        cv.close()
    except Exception as e:
        raise Exception(f"Erro ao converter o PDF: {str(e)}")
    
    return output_path

# Aplicativo Streamlit
st.title("Conversor PDF para DOCX")

uploaded_file = st.file_uploader("Escolha um arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    # Salva o arquivo PDF carregado no diretório temporário
    with open("uploaded_file.pdf", "wb") as f:
        f.write(uploaded_file.getbuffer())
    
    try:
        # Faz a conversão
        output_file = convert_pdf_to_docx("uploaded_file.pdf")
        st.success("Arquivo convertido com sucesso!")

        # Link para download
        with open(output_file, "rb") as f:
            st.download_button("Baixar DOCX", f, file_name=output_file)
    
    except Exception as e:
        st.error(f"Ocorreu um erro: {str(e)}")
