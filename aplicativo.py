import io
import os
import re
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
# Configurações da Página
# ==========================================
st.set_page_config(
    page_title="Conversor PDF p/ Docx", page_icon="📄", layout="centered"
)

# Logo do Web App (150px) e CSS para Centralização Completa da UI
logo_url = "https://i.imgur.com/VNPhtmN.jpeg"
st.markdown(
    f"""
    <style>
        /* Centralização da logo */
        .centered-logo {{
            display: flex;
            justify-content: center;
            margin-bottom: 15px;
        }}
        
        /* Centralização do rótulo e do grupo de radio buttons */
        div[data-testid="stRadio"] {{
            display: flex;
            flex-direction: column;
            align-items: center;
            text-align: center;
            width: 100%;
        }}
        
        div[data-testid="stRadio"] > label {{
            text-align: center;
            width: 100%;
            display: block;
        }}
        
        div[data-testid="stRadio"] [role="radiogroup"] {{
            display: inline-flex;
            flex-direction: column;
            align-items: flex-start;
            margin: 0 auto;
        }}
    </style>
    <div class="centered-logo">
        <img src="{logo_url}" width="150">
    </div>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# MÓDULO 1: Cópia Fiel (Para Impressão)
# ==========================================
def modo_copia_fiel(pdf_file):
    """Gera PDF intermediário com OCR e aplica pdf2docx para manter 100% da fidelidade visual."""
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
# MÓDULO 2: Texto Editável (Limpeza de Artefatos)
# ==========================================
def limpar_linha(texto_linha):
    """Filtra lixo de navegação e remove caracteres/números parasitas que aparecem antes das frases devido a ícones."""
    if not texto_linha:
        return None

    l_lower = texto_linha.lower().strip()

    blacklist = ["firefox", "about:blank", "lofl", "1 of 1"]
    if any(termo in l_lower for termo in blacklist):
        return None

    # Remove marcadores e caracteres estranhos
    texto_limpo = re.sub(
        r"^\s*[\(\[\{]?\d+[\)\]\}]?\s+(?=[A-Za-zÀ-ÿ])", "", texto_linha
    )
    texto_limpo = re.sub(
        r"^\s*[\(\[\{]?[A-Za-z0-9]{1,2}[\)\]\}]?\s+(?=[A-Za-zÀ-ÿ])",
        "",
        texto_limpo,
    )
    texto_limpo = re.sub(r"^\s*[^a-zA-Z0-9À-ÿ]+\s*(?=[A-Za-zÀ-ÿ])", "", texto_limpo)

    return texto_limpo.strip() if texto_limpo.strip() else None


def modo_texto_editavel(pdf_file):
    """Extrai texto e coordenadas via OCR, limpa artefatos e constrói parágrafos nativos no Word."""
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
            texto_bruto = " ".join(item["palavras"])
            texto_linha = limpar_linha(texto_bruto)

            if not texto_linha:
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
# Interface do Usuário (Streamlit UI)
# ==========================================
def main():
    # Título e Subtítulo Centralizados
    st.markdown(
        "<h2 style='text-align: center; font-size: 1.6rem; margin-top: 0;'>Conversor PDF p/ Docx</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; color: #555555; margin-bottom: 20px; font-size: 0.95rem;'>"
        "Envie seu arquivo PDF para convertê-lo em um arquivo Docx."
        "</p>",
        unsafe_allow_html=True,
    )

    # Opções do Modo de Conversão (Centralizadas)
    modo = st.radio(
        "Escolha o modo de conversão:",
        options=[
            "Cópia Fiel (Para Impressão - Mantém o layout 100% igual)",
            "Texto Editável (Para Edição - Texto limpo sem imagens de fundo)",
        ],
        index=0,
    )

    # Texto Instrucional Centralizado
    st.markdown(
        "<p style='text-align: center; margin-top: 15px; margin-bottom: 5px;'>Escolha seu arquivo abaixo:</p>",
        unsafe_allow_html=True,
    )

    # Área de Upload Reduzida em 50% e Centralizada (Proporção de colunas 1:2:1)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        pdf_file = st.file_uploader(
            "Escolha seu arquivo abaixo:",
            type="pdf",
            label_visibility="collapsed",
        )

    if pdf_file:
        with st.spinner("Convertendo arquivo..."):
            try:
                pdf_file.seek(0)

                if "Cópia Fiel" in modo:
                    docx_bytes = modo_copia_fiel(pdf_file)
                    nome_sufixo = "copia_fiel"
                else:
                    docx_bytes = modo_texto_editavel(pdf_file)
                    nome_sufixo = "editavel"

                st.success(
                    "Arquivo convertido com sucesso! Você pode baixar o arquivo Docx abaixo."
                )

                st.download_button(
                    label="Baixar Docx",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_{nome_sufixo}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Erro ao converter o arquivo: {e}")


if __name__ == "__main__":
    main()

# Rodapé Centralizado
st.markdown("""
---
<div style="text-align: center;">
    <h4 style="margin-bottom: 5px;">Web App - Conversor PDF p/ Docx</h4>
    <p style="color: #666666;">💬 Por Ary Ribeiro. Contato, através do email: <a href="mailto:aryribeiro@gmail.com">aryribeiro@gmail.com</a></p>
</div>
""", unsafe_allow_html=True)