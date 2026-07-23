import io
import os
import tempfile
import fitz  # PyMuPDF
from pdf2docx import Converter
from PIL import Image
from pypdf import PdfWriter
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


def pipeline_hibrido_pdf_para_docx(pdf_file):
    """Pipeline Híbrido:

    1. Renderiza páginas em 300 DPI e aplica OCR com coordenadas (X, Y).
    2. Gera um PDF intermediário com camada de texto estruturada.
    3. Alimenta o pdf2docx para recriar o layout, imagens, tabelas e fontes no Word.
    """
    pdf_bytes = pdf_file.read()
    doc_original = fitz.open(stream=pdf_bytes, filetype="pdf")
    merger = PdfWriter()

    # -------------------------------------------------------------
    # ETAPA 1: OCR por página para enriquecer com coordenadas
    # -------------------------------------------------------------
    for page_num in range(len(doc_original)):
        page = doc_original[page_num]

        # Renderiza a página em alta definição (300 DPI)
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes()))

        # O Tesseract gera um PDF de 1 página contendo o texto e suas coordenadas visuais
        pdf_ocr_bytes = pytesseract.image_to_pdf_or_hocr(
            img, extension="pdf", lang="por"
        )

        # Adiciona a página enriquecida ao acumulador
        page_pdf_file = io.BytesIO(pdf_ocr_bytes)
        merger.append(page_pdf_file)

    # Salva o PDF enriquecido temporariamente no disco
    with tempfile.NamedTemporaryFile(
        delete=False, suffix="_ocr.pdf"
    ) as temp_ocr_pdf:
        merger.write(temp_ocr_pdf)
        temp_ocr_pdf_path = temp_ocr_pdf.name

    temp_docx_path = temp_ocr_pdf_path.replace(".pdf", ".docx")

    # -------------------------------------------------------------
    # ETAPA 2: Reconstrução de Layout com pdf2docx
    # -------------------------------------------------------------
    try:
        cv = Converter(temp_ocr_pdf_path)
        # Processa todas as páginas recriando o layout com base nas coordenadas do OCR
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()

        return docx_bytes

    finally:
        # Limpeza de arquivos temporários do servidor
        if os.path.exists(temp_ocr_pdf_path):
            os.remove(temp_ocr_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)


# =========================
# Interface do App
# =========================
def main():
    st.title("Conversor Inteligente PDF p/ Docx 🚀")
    st.write(
        "Upload de arquivos com processamento em pipeline híbrido (OCR + Preservação de Layout)."
    )

    pdf_file = st.file_uploader("Escolha seu arquivo PDF abaixo:", type="pdf")

    if pdf_file:
        with st.spinner(
            "Executando Pipeline Híbrido (Visão Computacional + Reconstrução de Layout)..."
        ):
            try:
                pdf_file.seek(0)
                docx_bytes = pipeline_hibrido_pdf_para_docx(pdf_file)

                st.success("✨ Documento convertido com sucesso!")

                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro durante o pipeline de conversão: {e}")


if __name__ == "__main__":
    main()

# Rodapé
st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro. Contato, através do email: aryribeiro@gmail.com
""")