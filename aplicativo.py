import os
import tempfile
import streamlit as st
from pdf2docx import Converter

# =========================
# Configurações da Página
# =========================
st.set_page_config(
    page_title="Conversor PDF p/ Docx",
    page_icon="📄",
    layout="centered"
)

# Logo do Web App
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

def converter_pdf_para_docx_fiel(pdf_file):
    """
    Converte PDF para DOCX preservando ao máximo:
    - Tipografia (fontes, tamanhos, cores, negrito/itálico)
    - Espaçamento entre linhas e margens
    - Extração de imagens e logotipos originais
    - Tabelas e alinhamentos
    """
    # Criação dos arquivos temporários
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
        temp_pdf.write(pdf_file.read())
        temp_pdf_path = temp_pdf.name

    temp_docx_path = temp_pdf_path.replace(".pdf", ".docx")

    try:
        # Inicializa o conversor de alta fidelidade
        cv = Converter(temp_pdf_path)
        
        # A conversão padrão do pdf2docx extrai imagens automaticamente
        # e reconstrói parágrafos, fontes e espaçamentos
        cv.convert(temp_docx_path, start=0, end=None)
        cv.close()

        # Leitura dos bytes do DOCX gerado
        with open(temp_docx_path, "rb") as f:
            docx_bytes = f.read()

        return docx_bytes

    finally:
        # Limpeza rigorosa do sistema de arquivos no servidor
        if os.path.exists(temp_pdf_path):
            os.remove(temp_pdf_path)
        if os.path.exists(temp_docx_path):
            os.remove(temp_docx_path)


def main():
    st.title("Conversor PDF p/ Docx 📄")
    st.write(
        "Envie seu arquivo PDF. O sistema tentará preservar ao máximo a formatação, "
        "fontes, tamanhos, espaçamentos e imagens originais no arquivo Word."
    )

    pdf_file = st.file_uploader("Escolha seu arquivo PDF abaixo:", type="pdf")

    if pdf_file:
        with st.spinner("Analisando estrutura, fontes e imagens do PDF..."):
            try:
                pdf_file.seek(0)
                docx_bytes = converter_pdf_para_docx_fiel(pdf_file)

                st.success("✨ Conversão concluída! Clique no botão abaixo para baixar.")

                st.download_button(
                    label="📥 Baixar Documento (.docx)",
                    data=docx_bytes,
                    file_name=f"{pdf_file.name.replace('.pdf', '')}_convertido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
            except Exception as e:
                st.error(f"Erro no processamento da formatação: {e}")

if __name__ == "__main__":
    main()

# Rodapé
st.markdown("""
---
#### Web App - Conversor PDF p/ Docx
💬 Por Ary Ribeiro. Contato, através do email: aryribeiro@gmail.com
""")