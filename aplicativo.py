import streamlit as st
import pdf2docx

def convert_pdf_to_docx(pdf_file):
    output_path = "converted_file.docx"  # Defina o nome do arquivo de saída
    pdf2docx.convert(pdf_file, output_path)
    return output_path

st.title("Conversor de PDF para DOCX")

pdf_file = st.file_uploader("Carregar um arquivo PDF", type=["pdf"])

if pdf_file:
    st.write("Converting...")

    # Converte o arquivo
    output_file = convert_pdf_to_docx(pdf_file)

    # Exibe o link para download do arquivo convertido
    with open(output_file, "rb") as f:
        st.download_button(
            label="Baixar arquivo DOCX convertido",
            data=f,
            file_name=output_file,  # Define o nome correto do arquivo
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
