import streamlit as st
from pdf2docx import Converter
from io import BytesIO

def pdf_to_docx(pdf_file):
    """Converte um arquivo PDF para DOCX com a manutenção da formatação."""
    # Criando um conversor pdf2docx
    cv = Converter(pdf_file)
    
    # Criando um arquivo DOCX em memória
    docx_buffer = BytesIO()
    cv.convert(docx_buffer, start=0, end=None)  # Converte o PDF para o arquivo DOCX
    cv.close()
    
    docx_buffer.seek(0)  # Retorna ao início do arquivo em memória
    return docx_buffer

def main():
    """Interface do usuário para conversão de PDF para DOCX usando Streamlit."""
    st.title("Conversor de PDF para DOCX")

    st.write("Envie seu arquivo PDF para convertê-lo em um arquivo DOCX.")

    pdf_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

    if pdf_file:
        # Convertendo o PDF para DOCX em memória
        st.write("Convertendo PDF para DOCX...")
        docx_buffer = pdf_to_docx(pdf_file)

        if docx_buffer:
            st.success(f"Arquivo convertido com sucesso! Você pode baixar o arquivo DOCX abaixo.")
            
            # Botão de download com o arquivo DOCX gerado em memória
            st.download_button(
                label="Baixar DOCX",
                data=docx_buffer,
                file_name="output.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Erro ao converter o PDF. Tente novamente.")

if __name__ == "__main__":
    main()
