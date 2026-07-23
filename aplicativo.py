import os
import tempfile
import streamlit as st
from io import BytesIO
from pdf2docx import Converter
import fitz  # PyMuPDF (Já vem instalado com o pdf2docx)
from docx import Document

# =========================
# Configurações da Página
# =========================
st.set_page_config(page_title="Conversor PDF p/ Docx", page_icon="📄", layout="centered")

logo_url = "https://i.imgur.com/VNPhtmN.jpeg"
st.markdown(
    f"""
    <style>
        .centered-logo {{
            display: flex;
            justify-content: center;
            margin-bottom: 20px;
        }}
    </style>
    <div class="centered-logo">
        <img src="{logo_url}" width="300">
    </div>
    """,
    unsafe_allow_html=True,
)

# =========================
# Motores de Conversão
# =========================

def converte_layout_avancado(pdf_file):
    """Motor 1: Tenta preservar a geometria da página usando pdf2docx"""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_pdf_path = temp_pdf.name

    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    try:
        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        with open(temp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(temp_pdf_path): os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path): os.remove(temp_docx_path)

def converte_texto_limpo(pdf_file):
    """Motor 2: Ignora layout confuso e gráficos vetoriais. 
    Extrai apenas blocos de texto limpo de cima para baixo. Excelente para recibos de sites."""
    doc = Document()
    
    # Lendo o arquivo PDF em memória
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")
    
    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        # Extrai o dicionário de blocos da página
        blocks = page.get_text("blocks")
        
        # Ordena os blocos verticalmente (de cima para baixo na página)
        blocks.sort(key=lambda b: b[1])
        
        for b in blocks:
            text = b[4].strip()
            # Se for apenas lixo de quebra de linha ou vazio, ignora
            if text:
                doc.add_paragraph(text)
                
        # Adiciona quebra de página se não for a última
        if page_num < len(pdf_document) - 1:
            doc.add_page_break()
            
    # Salvar em bytes
    docx_buffer = BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()

# =========================
# Interface do App
# =========================
def main():
    st.title("Conversor PDF p/ Docx 🚀")
    st.write("Converta seus arquivos de forma inteligente. Selecione o tipo ideal de conversão abaixo:")

    # Painel de Controle Sênior: O Usuário escolhe a técnica
    modo_conversao = st.radio(
        "🛠️ **Escolha o Perfil do seu PDF:**",
        options=[
            "📄 Modo Padrão (Manter Layout, Tabelas e Imagens - Ideal para Textos/Livros)",
            "🧹 Modo Extração Limpa (Ignorar Layout - Ideal para Recibos, Boletos e Páginas da Internet)"
        ],
        index=0
    )

    pdf_file = st.file_uploader("Escolha seu arquivo abaixo (PDF):", type="pdf")

    if pdf_file:
        with st.spinner("Analisando e convertendo documento..."):
            try:
                # Resetar o ponteiro do arquivo para garantir leitura segura
                pdf_file.seek(0)
                
                if "Extração Limpa" in modo_conversao:
                    docx_bytes = converte_texto_limpo(pdf_file)
                else:
                    docx_bytes = converte_layout_avancado(pdf_file)

                st.success("✨ Arquivo convertido com sucesso! Pronto para download.")

                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Ocorreu um erro inesperado: {e}")

if __name__ == "__main__":
    main()

st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro: https://www.linkedin.com/in/aryribeiro
""")