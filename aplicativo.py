import base64
import hashlib
import logging
import os
import shutil
import signal
# subprocess: só para o conversor.py do próprio app, sem shell
import subprocess  # nosec B404
import sys
import tempfile
import threading
import time
import uuid
from collections import OrderedDict, deque
from pathlib import Path

import pymupdf as fitz  # aqui só para VALIDAR o PDF (sem renderizar nada)
import streamlit as st

# ==========================================
# Limites de proteção (memória, tempo e abuso)
# O contêiner do Streamlit Cloud tem ~1 GB. Medido em Docker: uma conversão
# de 2 páginas A4 no modo Cópia Fiel chega a ~430 MB (Python) + ~330 MB
# (Tesseract). Por isso: UMA conversão por vez, num processo separado, com
# teto de páginas, de tamanho de página, de tempo e de memória.
# ==========================================
APP_DIR = Path(__file__).resolve().parent
CONVERSOR = APP_DIR / "conversor.py"
LOGO_PATH = APP_DIR / "logo.jpeg"

MAX_UPLOAD_MB = 50             # também em .streamlit/config.toml (barra no servidor)
MAX_PAGINAS = 50
# a 300 dpi: A3 = 17,4 MP. Pico medido (1 CPU): A3 ~510 MB, A2 ~745 MB,
# 50 págs A4 ~580 MB. O piso de memória do Community Cloud é 690 MB.
MAX_MEGAPIXELS_PAGINA = 18
TIMEOUT_CONVERSAO_S = 600      # 50 págs A4 no modo Cópia Fiel: ~105 s com 1 CPU
LIMITE_MEMORIA_MB = 750        # soma do processo de conversão + Tesseract
CONVERSOES_SIMULTANEAS = 1
ESPERA_NA_FILA_S = 180
MAX_NA_FILA = 4
LIMITE_CONVERSOES = 10         # por sessão, dentro da janela abaixo
JANELA_LIMITE_S = 600
CACHE_MAX_MB = 64
CACHE_TTL_S = 3600
PREFIXO_TEMP = "pdfdoc_"

log = logging.getLogger("pdf_para_doc")
if not log.handlers:
    _h = logging.StreamHandler(sys.stderr)
    _h.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s %(message)s"))
    log.addHandler(_h)
    log.setLevel(logging.INFO)
    log.propagate = False

# ==========================================
# Configurações da Página
# ==========================================
st.set_page_config(
    page_title="Conversor PDF p/ Docx", page_icon="📄", layout="centered"
)


@st.cache_resource(show_spinner=False)
def logo_data_uri():
    """Logo servido pelo próprio app (antes vinha do imgur: mesma imagem,
    byte a byte). Sem dependência de terceiro para carregar a página."""
    try:
        return "data:image/jpeg;base64," + base64.b64encode(LOGO_PATH.read_bytes()).decode()
    except OSError:
        return "https://i.imgur.com/VNPhtmN.jpeg"


# Logo do Web App e CSS/JS para Centralização
logo_url = logo_data_uri()
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

    <div class="centered-logo">
        <img src="{logo_url}" width="150">
    </div>
    """,
    unsafe_allow_html=True,
)


# ==========================================
# Recursos compartilhados do processo (1 instância para todos os usuários)
# ==========================================
class Fila:
    """Semáforo de conversões + contador de quem espera (fila com teto)."""

    def __init__(self, vagas, max_espera):
        self.slots = threading.BoundedSemaphore(vagas)
        self.max_espera = max_espera
        self.esperando = 0
        self.lock = threading.Lock()

    def entrar_na_espera(self):
        with self.lock:
            if self.esperando >= self.max_espera:
                return False
            self.esperando += 1
            return True

    def sair_da_espera(self):
        with self.lock:
            self.esperando -= 1


class CacheResultados:
    """Cache em memória com TTL e teto em BYTES (não em entradas): um DOCX
    de 50 páginas pesa muito mais que um de 1. Chave = sha256 do PDF + modo,
    então converter o mesmo arquivo de novo devolve o mesmo resultado sem
    custo (idempotência)."""

    def __init__(self, max_bytes, ttl):
        self.max_bytes = max_bytes
        self.ttl = ttl
        self.itens = OrderedDict()
        self.total = 0
        self.lock = threading.Lock()

    def get(self, chave):
        with self.lock:
            item = self.itens.get(chave)
            if item is None:
                return None
            criado, dados = item
            if time.time() - criado > self.ttl:
                self._remover(chave)
                return None
            self.itens.move_to_end(chave)
            return dados

    def put(self, chave, dados):
        if len(dados) > self.max_bytes // 4:
            return
        with self.lock:
            if chave in self.itens:
                self._remover(chave)
            self.itens[chave] = (time.time(), dados)
            self.total += len(dados)
            while self.total > self.max_bytes and self.itens:
                self._remover(next(iter(self.itens)))

    def _remover(self, chave):
        _, dados = self.itens.pop(chave)
        self.total -= len(dados)


class LimiteDeUso:
    """Janela deslizante por cliente: N conversões a cada X segundos.
    Conversões servidas do cache não contam (não custam nada)."""

    def __init__(self, limite, janela, max_clientes=5000):
        self.limite = limite
        self.janela = janela
        self.max_clientes = max_clientes
        self.registros = {}
        self.lock = threading.Lock()

    def permitir(self, cliente):
        agora = time.time()
        with self.lock:
            if len(self.registros) > self.max_clientes:
                self._podar(agora)
            fila = self.registros.setdefault(cliente, deque())
            while fila and agora - fila[0] > self.janela:
                fila.popleft()
            if len(fila) >= self.limite:
                return False, int(self.janela - (agora - fila[0])) + 1
            fila.append(agora)
            return True, 0

    def _podar(self, agora):
        for c in [c for c, f in self.registros.items() if not f or agora - f[-1] > self.janela]:
            del self.registros[c]


@st.cache_resource(show_spinner=False)
def recursos():
    limpar_temporarios_orfaos()
    return (
        Fila(CONVERSOES_SIMULTANEAS, MAX_NA_FILA),
        CacheResultados(CACHE_MAX_MB * 1024 * 1024, CACHE_TTL_S),
        LimiteDeUso(LIMITE_CONVERSOES, JANELA_LIMITE_S),
    )


def limpar_temporarios_orfaos():
    """Remove pastas pdfdoc_* com mais de 1h (de execuções que morreram no
    meio) e PDFs/DOCX intermediários órfãos. Roda 1x por boot."""
    corte = time.time() - 3600
    tmp = Path(tempfile.gettempdir())
    for padrao in (PREFIXO_TEMP + "*", "*_ocr.pdf", "*_ocr.docx"):
        for p in tmp.glob(padrao):
            try:
                if p.stat().st_mtime < corte:
                    shutil.rmtree(p, ignore_errors=True) if p.is_dir() else p.unlink()
            except OSError:  # sumiu ou está em uso: tenta no próximo boot
                pass


def id_sessao():
    """Chave do limite de uso. Por que NÃO o IP: o proxy do Streamlit Cloud
    só entrega IPs internos no X-Forwarded-For (medido em 11/07/2026:
    "192.168.0.119, 10.13.215.3"), compartilhados por visitantes diferentes;
    um IP público ali, ou capturado pelo navegador (ipify), é escrito pelo
    próprio cliente e deixaria um atacante gastar o limite de outra pessoa."""
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return "sessao:" + get_script_run_ctx().session_id
    except Exception:  # noqa: BLE001
        return "desconhecido"


# ==========================================
# Validação do arquivo enviado (antes de gastar CPU/memória)
# ==========================================
class ArquivoRecusado(ValueError):
    """Mensagem já pronta para o usuário."""


def validar_pdf(dados):
    """Confere assinatura, integridade, senha, nº de páginas e tamanho de
    cada página (bomba de pixels). Devolve o nº de páginas."""
    if len(dados) > MAX_UPLOAD_MB * 1024 * 1024:
        raise ArquivoRecusado(f"Arquivo muito grande. O máximo é {MAX_UPLOAD_MB} MB.")
    if b"%PDF-" not in dados[:1024]:
        raise ArquivoRecusado("Este arquivo não é um PDF válido.")
    try:
        doc = fitz.open(stream=dados, filetype="pdf")
    except Exception:  # noqa: BLE001
        raise ArquivoRecusado("Não foi possível abrir este PDF: ele parece estar danificado.")
    try:
        # needs_pass só para senha de ABERTURA: PDF com restrição de permissão
        # (sem senha para abrir) converte normalmente, como sempre converteu.
        if doc.needs_pass:
            raise ArquivoRecusado("Este PDF está protegido por senha. Remova a senha e envie de novo.")
        paginas = doc.page_count
        if paginas == 0:
            raise ArquivoRecusado("Este PDF não tem páginas.")
        if paginas > MAX_PAGINAS:
            raise ArquivoRecusado(
                f"Este PDF tem {paginas} páginas. O máximo por conversão é {MAX_PAGINAS}: "
                "divida o arquivo em partes e converta uma de cada vez."
            )
        for i in range(paginas):
            r = doc[i].rect
            mp = (r.width / 72 * 300) * (r.height / 72 * 300) / 1e6
            if mp > MAX_MEGAPIXELS_PAGINA:
                raise ArquivoRecusado(
                    f"A página {i + 1} é grande demais para converter "
                    f"({r.width / 72 * 2.54:.0f} × {r.height / 72 * 2.54:.0f} cm). "
                    "O maior tamanho aceito é A3."
                )
        return paginas
    finally:
        doc.close()


# ==========================================
# Execução da conversão num processo separado (tempo e memória vigiados)
# ==========================================
class FalhaConversao(RuntimeError):
    def __init__(self, motivo):
        super().__init__(motivo)
        self.motivo = motivo  # "tempo" | "memoria" | "ocr" | "falha"


def _rss_do_grupo_mb(pgid):
    """Soma a memória residente de todos os processos do grupo (conversor +
    Tesseract). Só Linux (/proc); em outro SO devolve 0 e o teto não age."""
    total = 0
    pagina = os.sysconf("SC_PAGE_SIZE") if hasattr(os, "sysconf") else 4096
    proc = Path("/proc")
    if not proc.is_dir():
        return 0
    for d in proc.iterdir():
        if not d.name.isdigit():
            continue
        try:
            stat = (d / "stat").read_text()
            campos = stat[stat.rindex(")") + 2:].split()
            if int(campos[2]) != pgid:  # campo 5 do stat = pgrp
                continue
            total += int((d / "statm").read_text().split()[1]) * pagina
        except (OSError, ValueError, IndexError):
            continue
    return total // (1024 * 1024)


def _matar(proc):
    try:
        if os.name == "posix":
            os.killpg(proc.pid, signal.SIGKILL)  # leva o Tesseract junto
        else:
            proc.kill()
    except (ProcessLookupError, OSError):
        pass
    proc.wait()


def executar_conversor(modo, dados):
    pasta = Path(tempfile.gettempdir()) / f"{PREFIXO_TEMP}{uuid.uuid4().hex}"
    pasta.mkdir(parents=True)
    try:
        entrada, saida, erros = pasta / "entrada.pdf", pasta / "saida.docx", pasta / "erros.txt"
        entrada.write_bytes(dados)
        env = dict(os.environ, OMP_THREAD_LIMIT="1", TMPDIR=str(pasta))
        with open(erros, "wb") as ferr:
            # Sem shell; argumentos fixos: Python do app, conversor.py, modo de
            # uma lista fechada e caminhos gerados aqui (nada vem do usuário).
            proc = subprocess.Popen(  # nosec B603
                [sys.executable, str(CONVERSOR), modo, str(entrada), str(saida)],
                stdout=subprocess.DEVNULL, stderr=ferr, cwd=str(pasta), env=env,
                start_new_session=(os.name == "posix"),
            )
            inicio, pico = time.time(), 0
            while proc.poll() is None:
                time.sleep(0.25)
                if time.time() - inicio > TIMEOUT_CONVERSAO_S:
                    _matar(proc)
                    raise FalhaConversao("tempo")
                if os.name == "posix":
                    pico = max(pico, _rss_do_grupo_mb(proc.pid))
                    if pico > LIMITE_MEMORIA_MB:
                        _matar(proc)
                        raise FalhaConversao("memoria")
        if proc.returncode == 4:
            raise FalhaConversao("ocr")
        if proc.returncode != 0 or not saida.is_file():
            detalhe = erros.read_text(errors="replace")[-600:]
            log.error("conversor saiu com %s: %s", proc.returncode, detalhe)
            raise FalhaConversao("falha")
        log.info("convertido modo=%s em %.1fs pico=%sMB", modo, time.time() - inicio, pico)
        return saida.read_bytes()
    finally:
        shutil.rmtree(pasta, ignore_errors=True)


MENSAGENS_FALHA = {
    "tempo": "A conversão passou do tempo limite. Tente um PDF com menos páginas.",
    "memoria": "Este PDF exige mais memória do que o servidor tem disponível. "
               "Tente dividir o arquivo em partes menores.",
    "ocr": "O servidor está sobrecarregado no momento. Tente novamente em instantes.",
    "falha": "Não foi possível converter este arquivo. Ele pode estar danificado; "
             "tente novamente ou use outro PDF.",
}


def converter(dados, modo):
    """Valida, aplica cache, limite de uso e fila; devolve os bytes do DOCX
    ou None (a mensagem ao usuário já foi exibida)."""
    fila, cache, limite = recursos()
    chave = hashlib.sha256(dados).hexdigest() + ":" + modo
    em_cache = cache.get(chave)
    if em_cache is not None:
        return em_cache

    try:
        paginas = validar_pdf(dados)
    except ArquivoRecusado as e:
        st.error(str(e))
        return None

    # Limite por sessão (ver id_sessao). Memória, tempo e concorrência não
    # dependem de identificar o visitante: a fila protege o servidor.
    ok, espera = limite.permitir(id_sessao())
    if not ok:
        st.warning(f"Você atingiu o limite de conversões. Tente de novo em {max(1, espera // 60)} min.")
        log.warning("limite de uso atingido")
        return None

    # A vaga é devolvida num finally que cobre TAMBÉM a espera na fila: o
    # Streamlit pode interromper o script na saída do spinner (o usuário
    # mexeu na tela) logo depois de a vaga ter sido obtida.
    adquirido = fila.slots.acquire(blocking=False)
    try:
        if not adquirido:
            if not fila.entrar_na_espera():
                st.warning("O servidor está ocupado agora. Tente novamente em alguns minutos.")
                log.warning("fila cheia")
                return None
            try:
                with st.spinner("Na fila: outro arquivo está sendo convertido..."):
                    adquirido = fila.slots.acquire(timeout=ESPERA_NA_FILA_S)
            finally:
                fila.sair_da_espera()
            if not adquirido:
                st.warning("O servidor está ocupado agora. Tente novamente em alguns minutos.")
                return None

        log.info("convertendo modo=%s paginas=%s bytes=%s", modo, paginas, len(dados))
        with st.spinner("Convertendo arquivo..."):
            docx_bytes = executar_conversor(modo, dados)
            # Grava no cache ANTES de sair do spinner: se o usuário mexeu na
            # tela durante a conversão, o Streamlit interrompe o script na
            # próxima chamada st.* e o resultado seria jogado fora.
            cache.put(chave, docx_bytes)
    except FalhaConversao as e:
        log.warning("falha de conversao: %s", e.motivo)
        st.error(MENSAGENS_FALHA[e.motivo])
        return None
    except Exception:  # noqa: BLE001 — nunca mostrar detalhe interno ao usuário
        log.exception("erro inesperado na conversao")
        st.error(MENSAGENS_FALHA["falha"])
        return None
    finally:
        if adquirido:
            fila.slots.release()

    return docx_bytes


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
        "Envie PDF para converter em Docx"
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
                "Cópia Fiel (Para Impressão: layout idêntico)",
                "Texto Editável (Texto limpo sem imagens)",
            ],
            index=0,
        )

        # Upload de Arquivo
        pdf_file = st.file_uploader(
            " ",
            type="pdf",
            label_visibility="collapsed",
        )

        # Processamento, Alerta e Botão de Download dentro da coluna de 50%
        if pdf_file:
            if "Cópia Fiel" in modo:
                chave_modo, nome_sufixo = "fiel", "copia_fiel"
            else:
                chave_modo, nome_sufixo = "editavel", "editavel"

            docx_bytes = converter(pdf_file.getvalue(), chave_modo)
            if docx_bytes is None:
                return

            st.success(
                "Arquivo convertido com sucesso! Você pode baixar o arquivo Docx abaixo."
            )

            st.download_button(
                label="Baixar Docx",
                data=docx_bytes,
                file_name=f"{pdf_file.name.replace('.pdf', '')}_{nome_sufixo}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                width="stretch",
                on_click="ignore",  # baixar não reexecuta o app
            )


if __name__ == "__main__":
    main()

    st.markdown("""
<style>
    .main {
        background-color: #ffffff;
        color: #333333;
    }
    .block-container {
        padding-top: 1rem;
        padding-bottom: 0rem;
    }
    header {display: none !important;}
    footer {display: none !important;}
    #MainMenu {display: none !important;}
    div[data-testid="stAppViewBlockContainer"] {
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }
    div[data-testid="stVerticalBlock"] {
        gap: 0 !important;
        padding-top: 0 !important;
        padding-bottom: 0 !important;
    }
    .element-container {
        margin-top: 0 !important;
        margin-bottom: 0 !important;
    }
</style>
""", unsafe_allow_html=True)
