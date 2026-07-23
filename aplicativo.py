import io
import os
import tempfile
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from pdf2docx import Converter
from PIL import Image
from pypdf import PdfWriter
import pytesseract
import streamlit as st

# ==========================================
# Configurações da Página e Estilo
# ==========================================
st.set_page_config(
    page_title="Conversor PDF p/ Docx Profissional",
    page_icon="📄",
    layout="centered",
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


# ==========================================
# MÓDULO 1: Modo Cópia Fiel (Para Impressão)
# ==========================================
def modo_copia_fiel(pdf_file):
    """Pipeline Híbrido: Gera um PDF pesquisável com OCR e reconstrói o layout perfeito via pdf2docx.

    Excelente para impressão ou visualização idêntica.
    """
    pdf_bytes = pdf_file.read()
    doc_original = fitz.open(stream=pdf_bytes, filetype="pdf")
    merger = PdfWriter()

    for page_num in range(len(doc_original)):
        page = doc_original[page_num]
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes()))

        pdf_ocr_bytes = pytesseract.image_to_pdf_or_hocr(
            img, extension="pdf", lang="por"
        )
        merger.append(io.BytesIO(pdf_ocr_bytes))

    with tempfile.NamedTemporaryFile(
        delete=False, suffix="_ocr.pdf"
    ) as temp_ocr_pdf:
        merger.write(temp_ocr_pdf)
        temp_ocr_pdf_path = temp_ocr_pdf.name

    temp_docx_path = temp_ocr_pdf_path.replace(".pdf", ".docx")

    try:
        cv = Converter(temp_ocr_pdf_path)
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        with open(temp_docx_path, "rb") as f:
            return f.read()
    finally:
        if os.path.exists(temp_ocr_pdf_path):
            os.remove(temp_ocr_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)


# ==========================================
# MÓDULO 2: Modo Texto Editável (Sem Imagens de Fundo)
# ==========================================
def modo_texto_editavel(pdf_file):
    """Lê as caixas de OCR (image_to_data) e recria parágrafos nativos no Word.

    Livre de imagens travando a edição.
    """
    doc_word = Document()
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes()))

        data = pytesseract.image_to_data(
            img, lang="por", output_type=pytesseract.Output.DICT
        )

        linhas = {}
        num_items = len(data["text"])

        for i in range(num_items):
            texto = data["text"][i].strip()
            confianca = int(data["conf"][i])

            if confianca > 30 and texto:
                block_num = data["block_num"][i]
                line_num = data["line_num"][i]
                chave_linha = (block_num, line_num)

                top = data["top"][i]
                height = data["height"][i]

                if chave_linha not in linhas:
                    linhas[chave_linha] = {
                        "palavras": [],
                        "top": top,
                        "heights": [],
                    }

                linhas[chave_linha]["palavras"].append(texto)
                linhas[chave_linha]["heights"].append(height)

        linhas_ordenadas = sorted(linhas.values(), key=lambda x: x["top"])

        for item in linhas_ordenadas:
            texto_linha = " ".join(item["palavras"])
            l_lower = texto_linha.lower()

            # Remove ruídos de topo/rodapé de impressão de navegadores
            if any(
                termo in l_lower
                for termo in ["firefox", "about:blank", "1 of 1"]
            ):
                continue

            altura_media_px = sum(item["heights"]) / len(item["heights"])
            tamanho_fonte_pt = max(9, min(26, int(altura_media_px / 3.8)))

            p = doc_word.add_paragraph()
            run = p.add_run(texto_linha)
            run.font.name = "Arial"
            run.font.size = Pt(tamanho_fonte_pt)

            if (
                tamanho_fonte_pt >= 14
                or texto_linha.isdigit()
                or "Protocolo" in texto_linha
            ):
                run.bold = True

            p.paragraph_format.space_after = Pt(4)

        if page_num < len(pdf_document) - 1:
            doc_word.add_page_break()

    docx_buffer = io.BytesIO()
    doc_word.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()


# ==========================================
# MÓDULO 3: Modo Rápido (PDFs Nativos/Digitais)
# ==========================================
def modo_padrao_digital(pdf_file):
    """Conversão direta via pdf2docx para PDFs que já nasceram digitais (criados no Word, Canva, etc.)."""
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
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)


# ==========================================
# Interface do Usuário (Streamlit UI)
# ==========================================
def main():
    st.title("Conversor PDF p/ Docx 🚀")
    st.write(
        "Selecione o objetivo da sua conversão para obter o melhor resultado:"
    )

    # Seleção de Modo
    modo = st.radio(
        "🎯 **Escolha o Modo de Conversão:**",
        options=[
            "🎯 Cópia Fiel / Impressão (Fidelidade Visual 100% - Ideal para gerar réplicas perfeitas)",
            "✏️ Texto Editável (Livre de imagens de fundo - Ideal para alterar números, datas e textos)",
            "⚡ PDF Digital Padrão (Sem OCR - Mais rápido para arquivos gerados direto do Word/Canva)",
        ],
        index=0,
    )

    pdf_file = st.file_uploader("Escolha seu arquivo PDF abaixo:", type="pdf")

    if pdf_file:
        with st.spinner("Processando arquivo com a melhor tecnologia..."):
            try:
                pdf_file.seek(0)

                if "Cópia Fiel" in modo:
                    docx_bytes = modo_copia_fiel(pdf_file)
                    nome_sufixo = "replica_impressao"
                elif "Texto Editável" in modo:
                    docx_bytes = modo_texto_editavel(pdf_file)
                    nome_sufixo = "texto_editavel"
                else:
                    docx_bytes = modo_padrao_digital(pdf_file)
                    nome_sufixo = "convertido"

                st.success("✨ Conversão concluída com sucesso!")

                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_{nome_sufixo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro ao processar o arquivo: {e}")


if __name__ == "__main__":
    main()

# Rodapé
st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro. Contato, através do email: aryribeiro@gmail.com
""")