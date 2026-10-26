Obs.: caso o app esteja no modo "sleeping" (dormindo) ao entrar, basta clicar no botão que estará disponível e aguardar, para ativar o mesmo.
![logo-pdf-docx](https://github.com/user-attachments/assets/110a42c1-4866-4ba8-a292-75ff716076e9)

## Conversor PDF para DOCX

App no ar: https://pdf-para-doc.streamlit.app/

Web app em **Streamlit** que converte PDF em DOCX (Word), em dois modos:

- **Cópia Fiel (Para Impressão)**: cada página passa por OCR (Tesseract, português) e vira um PDF pesquisável, que o **pdf2docx** transforma em DOCX com o layout da página.
- **Texto Editável**: o OCR lê as linhas de cada página e o **python-docx** monta parágrafos nativos do Word, sem imagens.

## Limites e proteções

O servidor gratuito do Streamlit Community Cloud tem pouca memória. Para o app não cair:

- **Uma conversão por vez**, com fila de espera (até 4 pessoas). Quem passa disso recebe "servidor ocupado".
- A conversão roda num **processo separado** (`conversor.py`), com teto de **10 minutos** e de **750 MB** de memória. Se passar disso, o processo é encerrado e a pessoa recebe um aviso. O app continua no ar e libera toda a memória da conversão ao final.
- O arquivo é conferido antes de converter. Recusa: arquivo que não é PDF, PDF danificado ou com senha, mais de **50 páginas**, página maior que **A3** e arquivo acima de **50 MB**.
- Cada visitante (IP) pode fazer até **10 conversões a cada 10 minutos**.
- O mesmo arquivo no mesmo modo é convertido uma vez só e fica guardado por 1 hora, em até 64 MB de memória.
- Nenhum detalhe interno de erro aparece na tela. Os erros vão para o log do servidor, sem o nome do arquivo.

## Tecnologias

Streamlit, PyMuPDF, pdf2docx, python-docx, pytesseract (Tesseract OCR), Pillow, pypdf. As versões estão fixadas no `requirements.txt` e são atualizadas pelo Dependabot.

## Instalação local

1. Clone o repositório:
    ```bash
    git clone https://github.com/aryribeiro/pdf-para-doc.git
    ```
2. Instale o Tesseract com o idioma português (no Debian/Ubuntu: `apt install tesseract-ocr tesseract-ocr-por`).
3. Instale as dependências:
    ```bash
    pip install -r requirements.txt
    ```
4. Execute:
    ```bash
    streamlit run aplicativo.py
    ```

Para implantar no Streamlit Community Cloud, escolha **Python 3.13** em *Advanced settings*.

## Testes

A bancada roda num contêiner que imita o Streamlit Cloud (Debian, Python 3.13, Tesseract e 1 GB de memória). Ela confere que o DOCX gerado é idêntico, parte por parte, ao da versão anterior às proteções. Também testa a validação, o cache, o limite de uso e os tetos de tempo e de memória:

```bash
docker build -f tests/Dockerfile.smoke -t pdf-doc-smoke .
docker run --rm --memory=1g pdf-doc-smoke
```

O GitHub Actions roda a mesma bancada, o `pip-audit` (vulnerabilidades nas dependências) e o `bandit` (análise do código) a cada push e toda segunda-feira.

## Contato

💬 Por Ary Ribeiro. Para mais informações ou dúvidas, entre em contato via e-mail: [aryribeiro@gmail.com](mailto:aryribeiro@gmail.com)
