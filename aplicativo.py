import io
import os
import re
import tempfile
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches
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
        .centered-logo {{
            display: flex;
            justify-content: center;
            margin-bottom: 15px;
        }}
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
        from pdf2docx import Converter
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
    if not texto_linha:
        return None

    l_lower = texto_linha.lower().strip()
    blacklist = ["firefox", "about:blank", "lofl", "1 of 1"]
    if any(termo in l_lower for termo in blacklist):
        return None

    texto_limpo = re.sub(r"^\s*[\(\[\{]?\d+[\)\]\}]?\s+(?=[A-Za-zÀ-ÿ])", "", texto_linha)
    texto_limpo = re.sub(r"^\s*[\(\[\{]?[A-Za-z0-9]{1,2}[\)\]\}]?\s+(?=[A-Za-zÀ-ÿ])", "", texto_limpo)
    texto_limpo = re.sub(r"^\s*[^a-zA-Z0-9À-ÿ]+\s*(?=[A-Za-zÀ-ÿ])", "", texto_limpo)

    return texto_limpo.strip() if texto_limpo.strip() else None


def modo_texto_editavel(pdf_file):
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
                    linhas[chave_linha] = {"palavras": [], "top": top, "heights": []}

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

            if tamanho_fonte_pt >= 14 or texto_linha.isdigit() or "Protocolo" in texto_linha:
                run.bold = True

            p.paragraph_format.space_after = Pt(4)

        if page_num < len(pdf_document) - 1:
            doc_word.add_page_break()

    docx_buffer = io.BytesIO()
    doc_word.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()

import io
import os
import re
import tempfile
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
import docx.oxml as oxml
import docx.opc.constants as opc
from PIL import Image
from pypdf import PdfWriter
import pytesseract
import streamlit as st

# Helper para criar links clicáveis nativos no Word
def adicionar_link_clicavel(paragraph, url, text, color="0000FF", underline=True):
    part = paragraph.part
    r_id = part.relate_to(url, opc.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    hyperlink = oxml.parse_xml(
        f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" r:id="{r_id}"/>'
    )
    new_run = oxml.parse_xml(
        f'<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    rPr = oxml.parse_xml(
        f'<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    
    if color:
        c = oxml.parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = oxml.parse_xml('<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>')
        rPr.append(u)
        
    new_run.append(rPr)
    text_node = oxml.parse_xml(f'<w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">{text}</w:t>')
    new_run.append(text_node)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


# ==========================================
# MÓDULO 3: Fiel e Editável (Com Títulos Pretos, Links e Linhas)
# ==========================================
def modo_fiel_editavel(pdf_file):
    pdf_bytes = pdf_file.read()
    doc_word = Document()
    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_idx in range(len(pdf_doc)):
        page = pdf_doc[page_idx]
        
        # 1. Extração de Links
        links_pagina = page.get_links()
        
        # 2. Extração de Linhas/Vetores
        desenhos = page.get_drawings()
        linhas_y = []
        for d in desenhos:
            for item in d.get("items", []):
                if item[0] in ("l", "r"):  # Linhas ou Retângulos
                    rect = item[1] if item[0] == "r" else fitz.Rect(item[1], item[2])
                    # Verifica se é uma linha horizontal longa
                    if rect.width > 100 and rect.height < 5:
                        linhas_y.append(rect.y0)
        linhas_y.sort()

        # 3. Extração de Texto por Blocos
        blocks = page.get_text("dict", flags=fitz.TEXT_DEHYPHENATE)["blocks"]

        for b in blocks:
            if b.get("type") != 0:
                continue

            lines = b["lines"]
            p_atual = None
            ultimo_y1 = None
            ultimo_tamanho_fonte = None

            for l in lines:
                texto_linha = ""
                e_bold = False
                e_italic = False
                tamanho_max = 0.0
                bbox_linha = l["bbox"]

                for span in l["spans"]:
                    t = span["text"]
                    texto_linha += t
                    fonte_nome = span["font"].lower()
                    
                    if any(k in fonte_nome for k in ["bold", "black", "heavy"]):
                        e_bold = True
                    if any(k in fonte_nome for k in ["italic", "oblique"]):
                        e_italic = True
                    if span["size"] > tamanho_max:
                        tamanho_max = span["size"]

                line_str = texto_linha.strip()
                if not line_str:
                    continue

                line_str = re.sub(r"^(\d+\.)([A-Za-zÀ-ÿ])", r"\1 \2", line_str)

                y0, y1 = bbox_linha[1], bbox_linha[3]

                # Desenha linha divisória se houver um vetor próximo no Y
                if linhas_y and any(abs(y0 - ly) < 8 for ly in linhas_y):
                    p_linha = doc_word.add_paragraph()
                    p_linha.paragraph_format.space_before = Pt(6)
                    p_linha.paragraph_format.space_after = Pt(6)
                    p_linha_run = p_linha.add_run("―" * 45)
                    p_linha_run.font.color.rgb = RGBColor(180, 180, 180)

                # Identificação semântica
                e_bullet = line_str.startswith(("•", "-", "–", "*")) or bool(re.match(r"^\d+[\.\)]\s", line_str))
                e_titulo_principal = tamanho_max >= 15.0 or (e_bold and tamanho_max >= 13.0 and len(line_str) < 40)
                e_subtitulo = e_bold and (11.5 <= tamanho_max < 13.0) and len(line_str) < 50

                # Verifica se a linha possui link
                uri_link = None
                rect_linha = fitz.Rect(bbox_linha)
                for link in links_pagina:
                    if link.get("page") == page_idx or "uri" in link:
                        if rect_linha.intersects(link["from"]):
                            uri_link = link.get("uri")
                            break

                # Decisão de Parágrafo
                criar_novo_p = True
                if p_atual is not None and not e_bullet and not e_titulo_principal and not e_subtitulo:
                    distancia_vertical = y0 - (ultimo_y1 if ultimo_y1 is not None else y0)
                    if distancia_vertical < (tamanho_max * 1.4) and abs(tamanho_max - (ultimo_tamanho_fonte or tamanho_max)) < 2.0:
                        criar_novo_p = False

                if criar_novo_p:
                    if e_titulo_principal:
                        p_atual = doc_word.add_paragraph(style='Heading 1')
                        p_atual.paragraph_format.space_before = Pt(12)
                        p_atual.paragraph_format.space_after = Pt(4)
                    elif e_subtitulo:
                        p_atual = doc_word.add_paragraph(style='Heading 2')
                        p_atual.paragraph_format.space_before = Pt(8)
                        p_atual.paragraph_format.space_after = Pt(3)
                    else:
                        p_atual = doc_word.add_paragraph()
                        fmt = p_atual.paragraph_format
                        fmt.space_before = Pt(2)
                        fmt.space_after = Pt(2)
                        if e_bullet:
                            fmt.left_indent = Inches(0.25)

                if uri_link:
                    adicionar_link_clicavel(p_atual, uri_link, line_str)
                else:
                    run = p_atual.add_run(line_str if criar_novo_p else " " + line_str)
                    run.font.name = "Arial"
                    run.font.size = Pt(max(9, min(24, int(tamanho_max))))
                    
                    # Força a cor PRETA para todos os títulos e textos normais
                    run.font.color.rgb = RGBColor(0, 0, 0)
                    run.bold = e_bold
                    run.italic = e_italic

                ultimo_y1 = y1
                ultimo_tamanho_fonte = tamanho_max

        if page_idx < len(pdf_doc) - 1:
            doc_word.add_page_break()

    docx_buffer = io.BytesIO()
    doc_word.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()

# ==========================================
# Interface do Usuário (Streamlit UI)
# ==========================================
def main():
    st.markdown(
        "<h2 style='text-align: center; font-size: 1.6rem; margin-top: 0;'>Conversor PDF p/ Docx</h2>",
        unsafe_allow_html=True,
    )
    st.markdown(
        "<p style='text-align: center; color: #555555; margin-bottom: 20px; font-size: 0.95rem;'>"
        "Envie PDF para converter em Docx"
        "</p>",
        unsafe_allow_html=True,
    )

    col1, col2, col3 = st.columns([1, 2, 1])

    with col2:
        modo = st.radio(
            "Escolha o modo de conversão:",
            options=[
                "Cópia Fiel (Para Impressão: layout idêntico)",
                "Texto Editável (Texto limpo sem imagens)",
                "Fiel e Editável (Layout original e Textos Nativos)"
            ],
            index=2,
        )

        pdf_file = st.file_uploader(
            " ",
            type="pdf",
            label_visibility="collapsed",
        )

        if pdf_file:
            with st.spinner("Analisando estrutura e convertendo documento..."):
                try:
                    pdf_file.seek(0)

                    if "Cópia Fiel (Para Impressão" in modo:
                        docx_bytes = modo_copia_fiel(pdf_file)
                        nome_sufixo = "copia_fiel"
                    elif "Texto Editável" in modo:
                        docx_bytes = modo_texto_editavel(pdf_file)
                        nome_sufixo = "editavel"
                    else:
                        docx_bytes = modo_fiel_editavel(pdf_file)
                        nome_sufixo = "fiel_editavel_nativo"

                    st.success("Arquivo convertido com sucesso!")

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

    st.markdown("""
<style>
    .main {background-color: #ffffff; color: #333333;}
    .block-container {padding-top: 1rem; padding-bottom: 0rem;}
    header {display: none !important;}
    footer {display: none !important;}
    #MainMenu {display: none !important;}
    div[data-testid="stAppViewBlockContainer"] {padding-top: 0 !important; padding-bottom: 0 !important;}
    div[data-testid="stVerticalBlock"] {gap: 0 !important; padding-top: 0 !important; padding-bottom: 0 !important;}
    .element-container {margin-top: 0 !important; margin-bottom: 0 !important;}
</style>
""", unsafe_allow_html=True)