import streamlit as st
from pdf2docx import Converter
import tempfile
import os

# =========================
# Configurações da Página (Opcional, mas recomendado)
# =========================
st.set_page_config(page_title="Conversor PDF p/ Docx", page_icon="📄")

# =========================
# Logo do web app
# =========================

# URL do logo
logo_url = "https://i.imgur.com/VNPhtmN.jpeg"

# Exibindo o logo pela URL[cite: 1]
# HTML + CSS para centralizar[cite: 1]
st.markdown(
    f"""
    <style>
        .centered-logo {{
            display: flex;
            justify-content: center;
        }}
    </style>
    <div class="centered-logo">
        <img src="{logo_url}" width="300">
    </div>
    """,
    unsafe_allow_html=True
)

def convert_pdf_to_docx_with_layout(pdf_file):
    """
    Salva o PDF temporariamente, converte preservando layout (tabelas, imagens, fontes)
    e retorna o arquivo DOCX em formato de bytes.
    """
    # 1. Criação de arquivos temporários
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_pdf_path = temp_pdf.name
        
    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")
    
    try:
        # 2. Conversão usando pdf2docx (reconstrói o layout)
        cv = Converter(temp_pdf_path)
        # start=0, end=None converte todas as páginas
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()
        
        # 3. Leitura do arquivo DOCX gerado
        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()
            
        return docx_bytes
    
    finally:
        # 4. Limpeza de memória do servidor (Muito importante em Cloud!)
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)

def main():
    """Interface do usuário para conversão de PDF para DOCX usando Streamlit."""
    st.title("Conversor PDF p/ Docx")

    st.write("Envie seu arquivo PDF para convertê-lo em um arquivo Docx mantendo o layout.")

    pdf_file = st.file_uploader("Escolha seu arquivo abaixo:", type="pdf")[cite: 1]

    if pdf_file:
        with st.spinner("Analisando layout e convertendo... Isso pode levar alguns segundos dependendo do tamanho do PDF."):
            try:
                docx_bytes = convert_pdf_to_docx_with_layout(pdf_file)
                
                st.success("Arquivo convertido com sucesso! Você pode baixar o arquivo Docx abaixo.")[cite: 1]
                
                # Botão de download com o arquivo DOCX gerado[cite: 1]
                st.download_button(
                    label="Baixar Docx",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"[cite: 1]
                )
            except Exception as e:
                st.error(f"Ocorreu um erro ao converter o arquivo: {e}")

if __name__ == "__main__":
    main()

# Rodapé com informações de contato (em vermelho)[cite: 1]
st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro: https://www.linkedin.com/in/aryribeiro
""")