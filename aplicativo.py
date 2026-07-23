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

# Logo do Web App e CSS/JS para Centralização
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
        
        /* Centralização e impedimento de quebra de linha no stRadio */
        div[data-testid="stRadio"] {{
            display: flex !important;
            flex-direction: column !important;
            align-items: center !important;
            justify-content: center !important;
            text-align: center !important;
            width: 100% !important;
        }}
        
        div[data-testid="stRadio"] > label,
        div[data-testid="stRadio"] [data-testid="stWidgetLabel"] {{
            text-align: center !important;
            width: 100% !important;
            display: flex !important;
            justify-content: center !important;
            white-space: nowrap !important;
        }}
        
        div[data-testid="stRadio"] [role="radiogroup"] {{
            display: inline-flex !important;
            flex-direction: column !important;
            align-items: flex-start !important;
            margin: 0 auto !important;
            width: max-content !important;
        }}

        /* Força os textos das opções a não quebrarem linha */
        div[data-testid="stRadio"] label p,
        div[data-testid="stRadio"] label span,
        div[data-testid="stRadio"] label {{
            white-space: nowrap !important;
            word-break: keep-all !important;
        }}
    </style>

    <script>
        function forcarLinhaUnicaECentralizar() {{
            const radioContainer = window.parent.document.querySelector('div[data-testid="stRadio"]');
            if (radioContainer) {{
                radioContainer.style.setProperty('display', 'flex', 'important');
                radioContainer.style.setProperty('flex-direction', 'column', 'important');
                radioContainer.style.setProperty('align-items', 'center', 'important');
                
                const labels = radioContainer.querySelectorAll('label, p, span');
                labels.forEach(el => {{
                    el.style.setProperty('white-space', 'nowrap', 'important');
                    el.style.setProperty('word-break', 'keep-all', 'important');
                }});
                
                const group = radioContainer.querySelector('div[role="radiogroup"]');
                if (group) {{
                    group.style.setProperty('margin', '0 auto', 'important');
                    group.style.setProperty('width', 'max-content', 'important');
                }}
            }}
        }}
        setTimeout(forcarLinhaUnicaECentralizar, 200);
        setTimeout(forcarLinhaUnicaECentralizar, 600);
    </script>

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
        "Envie PDF para convertê-lo em Docx."
        "</p>",
        unsafe_allow_html=True,
    )

    # Bloco Centralizado (50% de largura da tela)
    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        # Modo de Conversão
        modo = st.radio(
            "Escolha o modo de conversão:",
            options=[
                "Cópia Fiel (Para Impressão - Mantém o layout 100% igual)",
                "Texto Editável (Para Edição - Texto limpo sem imagens de fundo)",
            ],
            index=0,
        )

        # Instrução de Upload
        st.markdown(
            "<p style='text-align: center; margin-top: 15px; margin-bottom: 5px;'>Escolha seu arquivo abaixo:</p>",
            unsafe_allow_html=True,
        )

        # Upload de Arquivo
        pdf_file = st.file_uploader(
            "Escolha seu arquivo abaixo:",
            type="pdf",
            label_visibility="collapsed",
        )

        # Processamento, Alerta e Botão de Download dentro da coluna de 50%
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

<div style="text-align: center;">
    <p style="color: #666666;">💬 Por Ary Ribeiro: <a href="https://www.linkedin.com/in/aryribeiro">https://www.linkedin.com/in/aryribeiro</a></p>
</div>
""", unsafe_allow_html=True)