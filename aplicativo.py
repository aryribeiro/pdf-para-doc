import io
import os
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import pytesseract
import streamlit as st

# =========================
# Configurações da Página
# =========================
st.set_page_config(
    page_title="Conversor PDF p/ Docx Editável",
    page_icon="📄",
    layout="centered"
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


def converter_pdf_para_docx_editavel(pdf_file):
    """
    Lê o PDF, processa via OCR obtendo coordenadas de cada linha,
    e reconstrói o arquivo com parágrafos NATIVOS do Word (100% editáveis).
    """
    doc_word = Document()
    pdf_bytes = pdf_file.read()
    pdf_document = fitz.open(stream=pdf_bytes, filetype="pdf")

    for page_num in range(len(pdf_document)):
        page = pdf_document[page_num]

        # 1. Renderiza a página em 300 DPI para alta precisão
        pix = page.get_pixmap(dpi=300)
        img = Image.open(io.BytesIO(pix.tobytes()))

        # 2. Extrai dados estruturados do OCR (palavras, posições, alturas e blocos)
        data = pytesseract.image_to_data(img, lang="por", output_type=pytesseract.Output.DICT)

        # 3. Agrupa as palavras por linhas e blocos
        linhas = {}
        num_items = len(data["text"])

        for i in range(num_items):
            texto = data["text"][i].strip()
            confianca = int(data["conf"][i])

            # Filtra ruídos ou espaços em branco
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
                        "heights": []
                    }

                linhas[chave_linha]["palavras"].append(texto)
                linhas[chave_linha]["heights"].append(height)

        # 4. Ordena as linhas de cima para baixo
        linhas_ordenadas = sorted(linhas.values(), key=lambda x: x["top"])

        # 5. Adiciona cada linha como um parágrafo nativo e editável no Word
        for item in linhas_ordenadas:
            texto_linha = " ".join(item["palavras"])

            # Ignora cabeçalhos e rodapés indesejados de impressão do navegador
            l_lower = texto_linha.lower()
            if any(termo in l_lower for termo in ["firefox", "about:blank", "1 of 1", "09/07/2026, 09:26"]):
                continue

            # Calcula o tamanho da fonte em pontos (pt) baseado na altura dos pixels
            altura_media_px = sum(item["heights"]) / len(item["heights"])
            
            # Conversão aproximada: 300 DPI (1 pt ≈ 4.16 px)
            tamanho_fonte_pt = max(9, min(28, int(altura_media_px / 3.8)))

            p = doc_word.add_paragraph()
            run = p.add_run(texto_linha)
            run.font.name = "Arial"
            run.font.size = Pt(tamanho_fonte_pt)

            # Aplica Negrito automático para títulos e números em destaque
            if tamanho_fonte_pt >= 14 or texto_linha.isdigit() or "Protocolo" in texto_linha:
                run.bold = True

            # Ajusta o espaçamento entre linhas para manter a leitura agradável
            p.paragraph_format.space_after = Pt(4)

        if page_num < len(pdf_document) - 1:
            doc_word.add_page_break()

    # Salva o arquivo em memória
    docx_buffer = io.BytesIO()
    doc_word.save(docx_buffer)
    docx_buffer.seek(0)
    return docx_buffer.read()


# =========================
# Interface do App
# =========================
def main():
    st.title("Conversor PDF p/ Docx Editável 📝")
    st.write("Converta comprovantes, boletos e PDFs em documentos Word **100% editáveis**.")

    pdf_file = st.file_uploader("Escolha seu arquivo PDF abaixo:", type="pdf")

    if pdf_file:
        with st.spinner("Extraindo textos e gerando parágrafos editáveis no Word..."):
            try:
                pdf_file.seek(0)
                docx_bytes = converter_pdf_para_docx_editavel(pdf_file)

                st.success("✨ Arquivo convertido com sucesso! O texto agora é 100% editável.")

                st.download_button(
                    label="📥 Baixar Documento Editável (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_editavel.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro no processamento do arquivo: {e}")


if __name__ == "__main__":
    main()

# Rodapé
st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro. Contato, através do email: aryribeiro@gmail.com
""")