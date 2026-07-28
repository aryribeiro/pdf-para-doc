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


def limpa_nome_fonte(font_name):
    """Normaliza o nome da fonte removendo prefixos de subconjuntos de PDF (ex: ABCDEF+Arial -> Arial)"""
    if not font_name:
        return "Arial"
    limpo = re.sub(r'^[A-Z]{6}\+', '', font_name)
    limpo = limpo.split('-')[0].split(',')[0]
    return limpo if limpo else "Arial"


def adicionar_borda_inferior_xml(paragraph, color_hex="000000", sz="12"):
    """Injeta a linha preta horizontal vetorial como uma borda inferior nativa do Word (OpenXML)"""
    pPr = paragraph._p.get_or_add_pPr()
    pBdr = oxml.parse_xml(
        f'<w:pBdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:bottom w:val="single" w:sz="{sz}" w:space="1" w:color="{color_hex}"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)


def adicionar_link_clicavel(paragraph, url, text, font_name, font_size, e_bold, e_italic, rgb_color):
    """Cria um hiperlink OpenXML preservando rigorosamente a fonte, tamanho e estilo do texto original do PDF"""
    part = paragraph.part
    r_id = part.relate_to(url, opc.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    
    hyperlink = oxml.parse_xml(
        f'<w:hyperlink xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        f'r:id="{r_id}"/>'
    )
    new_run = oxml.parse_xml(
        '<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    rPr = oxml.parse_xml(
        '<w:rPr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'
    )
    
    # 1. Preserva a família exata da fonte
    nome_fonte = limpa_nome_fonte(font_name)
    f_elem = oxml.parse_xml(
        f'<w:rFonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        f'w:ascii="{nome_fonte}" w:hAnsi="{nome_fonte}"/>'
    )
    rPr.append(f_elem)

    # 2. Preserva a cor exata do texto
    hex_color = f"{rgb_color[0]:02X}{rgb_color[1]:02X}{rgb_color[2]:02X}"
    c = oxml.parse_xml(f'<w:color xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{hex_color}"/>')
    rPr.append(c)
    
    # 3. Adiciona o sublinhado de link
    u = oxml.parse_xml('<w:u xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="single"/>')
    rPr.append(u)
    
    # 4. Preserva o tamanho exato em meio-pontos (half-points)
    sz_val = str(int(font_size * 2))
    sz_elem = oxml.parse_xml(f'<w:sz xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" w:val="{sz_val}"/>')
    rPr.append(sz_elem)

    if e_bold:
        rPr.append(oxml.parse_xml('<w:b xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))
    if e_italic:
        rPr.append(oxml.parse_xml('<w:i xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>'))

    new_run.append(rPr)
    
    text_node = oxml.OxmlElement('w:t')
    text_node.text = text
    new_run.append(text_node)
    
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)


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


# ==========================================
# MÓDULO 3: Fiel e Editável (Precisão Estrita 1:1)
# ==========================================
def modo_fiel_editavel(pdf_file):
    """
    Motor 1:1 com Preservação de Formatação Absoluta:
    - Mapeia diretamente cada linha do PDF para uma linha no Word (sem unir nem dividir).
    - Captura e desenha linhas pretas horizontais divisórias como bordas OpenXML.
    - Mantém tamanhos de fonte, cores RGB e fontes de hiperlinks 100% fiéis ao PDF.
    """
    pdf_bytes = pdf_file.read()
    doc_word = Document()

    # Ajusta as margens padrão da página
    for section in doc_word.sections:
        section.top_margin = Inches(0.5)
        section.bottom_margin = Inches(0.5)
        section.left_margin = Inches(0.5)
        section.right_margin = Inches(0.5)

    pdf_doc = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_idx in range(len(pdf_doc)):
        page = pdf_doc[page_idx]
        links_pagina = page.get_links()

        elementos = []

        # 1. Extração de Linhas Horizontais Pretas Divisórias (Vetores)
        desenhos = page.get_drawings()
        for d in desenhos:
            for item in d.get("items", []):
                if item[0] in ("l", "r"):
                    rect = item[1] if item[0] == "r" else fitz.Rect(item[1], item[2])
                    if rect.width > 60 and rect.height <= 8:
                        color_tuple = d.get("color") or d.get("fill") or (0, 0, 0)
                        hex_color = f"{int(color_tuple[0]*255):02X}{int(color_tuple[1]*255):02X}{int(color_tuple[2]*255):02X}"
                        elementos.append({
                            "tipo": "vetor_linha",
                            "y0": rect.y0,
                            "y1": rect.y1,
                            "bbox": rect,
                            "color_hex": hex_color if hex_color != "FFFFFF" else "000000"
                        })

        # 2. Extração de Linhas de Texto (1 Linha do PDF = 1 Parágrafo no Word)
        blocks = page.get_text("dict", flags=fitz.TEXT_DEHYPHENATE)["blocks"]
        for b in blocks:
            if b.get("type") != 0:
                continue
            for line in b["lines"]:
                elementos.append({
                    "tipo": "texto",
                    "y0": line["bbox"][1],
                    "y1": line["bbox"][3],
                    "bbox": line["bbox"],
                    "spans": line["spans"]
                })

        # Ordena todos os elementos espacialmente da parte superior para a inferior
        elementos.sort(key=lambda item: item["y0"])

        ultimo_y1 = None

        for el in elementos:
            delta_y = (el["y0"] - ultimo_y1) if ultimo_y1 is not None else 0.0

            # Renderização de Linhas Pretas Divisórias
            if el["tipo"] == "vetor_linha":
                p_vetor = doc_word.add_paragraph()
                fmt = p_vetor.paragraph_format
                fmt.space_before = Pt(max(2.0, min(delta_y, 24.0)))
                fmt.space_after = Pt(2.0)
                adicionar_borda_inferior_xml(p_vetor, color_hex=el["color_hex"], sz="12")
                ultimo_y1 = el["y1"]
                continue

            # Renderização de Linhas de Texto
            line_spans = el["spans"]
            if not line_spans:
                continue

            p = doc_word.add_paragraph()
            fmt = p.paragraph_format
            
            if delta_y > 2.0 and delta_y < 100.0:
                fmt.space_before = Pt(round(delta_y, 1))
            else:
                fmt.space_before = Pt(0)

            fmt.space_after = Pt(0)
            fmt.line_spacing = 1.0

            rect_linha = fitz.Rect(el["bbox"])

            for span in line_spans:
                texto_span = span["text"]
                if not texto_span:
                    continue

                fonte_nome = span["font"]
                fonte_lower = fonte_nome.lower()
                e_bold = any(k in fonte_lower for k in ["bold", "black", "heavy", "w7", "w8", "w9"])
                e_italic = any(k in fonte_lower for k in ["italic", "oblique"])
                tamanho_exact_pt = span["size"]

                color_int = span.get("color", 0)
                r = (color_int >> 16) & 0xFF
                g = (color_int >> 8) & 0xFF
                b_color = color_int & 0xFF

                # Identifica se a região do span intersecta com um hiperlink
                bbox_span = fitz.Rect(span["bbox"])
                uri_link = None
                for link in links_pagina:
                    if link.get("page") == page_idx or "uri" in link:
                        if bbox_span.intersects(link["from"]) or rect_linha.intersects(link["from"]):
                            uri_link = link.get("uri")
                            break

                if uri_link:
                    adicionar_link_clicavel(
                        paragraph=p,
                        url=uri_link,
                        text=texto_span,
                        font_name=fonte_nome,
                        font_size=tamanho_exact_pt,
                        e_bold=e_bold,
                        e_italic=e_italic,
                        rgb_color=(r, g, b_color)
                    )
                else:
                    run = p.add_run(texto_span)
                    run.font.name = limpa_nome_fonte(fonte_nome)
                    run.font.size = Pt(tamanho_exact_pt)
                    run.font.color.rgb = RGBColor(r, g, b_color)
                    run.bold = e_bold
                    run.italic = e_italic

            ultimo_y1 = el["y1"]

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