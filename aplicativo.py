import io
import os
import tempfile
import fitz  # PyMuPDF
from docx import Document
from PIL import Image
import pytesseract
import streamlit as st

# =========================
# Configurações da Página
# =========================
st.set_page_config(
    page_title="Conversor PDF p/ Docx", page_icon="📄", layout="centered"
)

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
    """Motor 1: Preserva geometria, tabelas e estilos (Ideal para PDFs nativos do Word/PDFs normais)."""
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_pdf_path = temp_pdf.name

    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    try:
        from pdf2docx import Converter

        cv = Converter(temp_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        with open(temp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)


def converte_com_ocr(pdf_file):
    """Motor 2: Renderiza a página em imagem 300 DPI e aplica Visão Computacional (OCR)

    Le com precisao números e textos que foram 'desenhados' ou escaneados.
    """
    doc = Document()
    pdf_document = fitz.open(stream=pdf_file.read(), filetype="pdf")

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]

        # Renderiza a página como imagem em alta definição (300 DPI)
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes()))

        # Aplica o OCR reconhecendo Português
        text_ocr = pytesseract.image_to_string(img, lang="por")

        # Insere o texto extraído no Word
        if text_ocr.strip():
            for line in text_ocr.split("\n"):
                if line.strip():
                    doc.add_paragraph(line.strip())

        if page_num < len(pdf_document) - 1:
            doc.add_page_break()

    docx_buffer = io.BytesIO()
    doc.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()


# =========================
# Interface do App
# =========================


def main():
    st.title("Conversor PDF p/ Docx 🚀")
    st.write(
        "Converta seus arquivos com máxima precisão. Escolha a melhor estratégia abaixo:"
    )

    modo_conversao = st.radio(
        "🛠️ **Escolha o Perfil do seu PDF:**",
        options=[
            "📄 Modo Padrão (Preserva Layout, Tabelas e Formatação de Documentos Comuns)",
            "👁️ Modo OCR com IA (Para PDFs Escaneados, Boletos e Comprovantes da Web)",
        ],
        index=0,
    )

    pdf_file = st.file_uploader("Escolha seu arquivo abaixo (PDF):", type="pdf")

    if pdf_file:
        with st.spinner("Analisando e processando documento..."):
            try:
                pdf_file.seek(0)

                if "Modo OCR" in modo_conversao:
                    docx_bytes = converte_com_ocr(pdf_file)
                else:
                    docx_bytes = converte_layout_avancado(pdf_file)

                st.success("✨ Arquivo convertido com sucesso!")

                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Ocorreu um erro no processamento: {e}")


if __name__ == "__main__":
    main()

st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro. Contato, através do email: aryribeiro@gmail.com
""")