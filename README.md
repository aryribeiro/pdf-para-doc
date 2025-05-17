Obs.: caso o app esteja no modo "sleeping" (dormindo) ao entrar, basta clicar no botão que estará disponível e aguardar, para ativar o mesmo.
![logo-pdf-docx](https://github.com/user-attachments/assets/110a42c1-4866-4ba8-a292-75ff716076e9)

## Conversor PDF para DOCX

Este é um web app desenvolvido com **Streamlit** que permite converter arquivos PDF em arquivos DOCX. O app utiliza a biblioteca **PyPDF2** para extrair texto de PDFs e **python-docx** para gerar arquivos DOCX a partir do texto extraído.

## Funcionalidade

- Carregar um arquivo PDF.
- Extrair o texto do PDF.
- Converter o texto extraído em um arquivo DOCX.
- Baixar o arquivo DOCX gerado.

## Tecnologias Usadas

- **Streamlit**: Para construir a interface do usuário.
- **PyPDF2**: Para ler e extrair texto de arquivos PDF.
- **python-docx**: Para gerar documentos DOCX a partir do texto extraído.

## Instalação

1. Clone o repositório:
    ```bash
    git clone https://github.com/aryribeiro/pdf-para-doc.git
    ```

2. Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```

## Como Usar

1. Execute o aplicativo Streamlit:
    ```bash
    streamlit run app.py
    ```

2. Acesse o aplicativo na URL fornecida pelo Streamlit no seu terminal.
   
3. Envie um arquivo PDF para converter em DOCX.

## Contato

💬 Por Ary Ribeiro. Para mais informações ou dúvidas, entre em contato via e-mail: [aryribeiro@gmail.com](mailto:aryribeiro@gmail.com)