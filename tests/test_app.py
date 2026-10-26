"""Bancada do pdf-para-doc. Roda no espelho do Streamlit Cloud (tests/Dockerfile.smoke).

1. Fidelidade: o conversor.py gera o MESMO DOCX que o aplicativo original
   (checkpoint-pre-devsecops) — comparado parte a parte do ZIP.
2. Validação do upload, cache, limite de uso e os tetos de tempo e memória.
"""
import io
import os
import sys
import time
import zipfile
from pathlib import Path

import pytest

RAIZ = Path(__file__).resolve().parents[1]
FIX = RAIZ / "tests" / "fixtures"
BASE = RAIZ / "tests" / "baseline_out"
sys.path.insert(0, str(RAIZ))

import aplicativo as app  # noqa: E402


def partes_docx(dados):
    """Conteúdo do DOCX sem docProps/core.xml (carimbo de data/hora)."""
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        return {n: z.read(n) for n in z.namelist() if n != "docProps/core.xml"}


# ---------------------------------------------------------------- fidelidade
@pytest.mark.parametrize("pdf", ["digital", "escaneado"])
@pytest.mark.parametrize("modo", ["fiel", "editavel"])
def test_docx_identico_ao_original(pdf, modo):
    gabarito = BASE / f"{pdf}_{modo}.docx"
    if not gabarito.is_file():
        pytest.skip("gabarito ausente: rode tests/baseline_original.py antes")
    novo = app.executar_conversor(modo, (FIX / f"{pdf}.pdf").read_bytes())
    esperado, obtido = partes_docx(gabarito.read_bytes()), partes_docx(novo)
    assert sorted(obtido) == sorted(esperado)
    diferentes = [n for n in esperado if esperado[n] != obtido[n]]
    assert diferentes == [], f"partes diferentes do original: {diferentes}"


# ---------------------------------------------------------------- validação
def test_aceita_pdf_normal():
    assert app.validar_pdf((FIX / "digital.pdf").read_bytes()) == 2


@pytest.mark.parametrize("arquivo, trecho", [
    ("nao_e_pdf.pdf", "não é um PDF"),
    ("senha.pdf", "senha"),
    ("muitas_paginas.pdf", "120 páginas"),
    ("pagina_gigante.pdf", "grande demais"),
])
def test_recusa(arquivo, trecho):
    with pytest.raises(app.ArquivoRecusado, match=trecho):
        app.validar_pdf((FIX / arquivo).read_bytes())


def test_recusa_pdf_truncado():
    dados = (FIX / "digital.pdf").read_bytes()
    with pytest.raises(app.ArquivoRecusado):
        app.validar_pdf(dados[:300])


def test_recusa_acima_do_teto_de_upload():
    with pytest.raises(app.ArquivoRecusado, match="muito grande"):
        app.validar_pdf(b"%PDF-1.7\n" + b"0" * (app.MAX_UPLOAD_MB * 1024 * 1024))


# ---------------------------------------------------------------- cache e limite
def test_cache_teto_em_bytes_e_lru():
    c = app.CacheResultados(max_bytes=100, ttl=60)
    c.put("a", b"x" * 20)
    c.put("b", b"x" * 20)
    c.get("a")                       # "a" fica mais recente
    for k in "cdef":
        c.put(k, b"x" * 20)
    assert c.total <= 100
    assert c.get("a") is not None and c.get("b") is None


def test_cache_ignora_item_gigante_e_expira():
    c = app.CacheResultados(max_bytes=100, ttl=0)
    c.put("grande", b"x" * 60)
    assert c.get("grande") is None
    c.put("p", b"x")
    time.sleep(0.01)
    assert c.get("p") is None and c.total == 0


def test_limite_de_uso_por_cliente():
    lim = app.LimiteDeUso(limite=3, janela=60)
    assert [lim.permitir("ip:1")[0] for _ in range(4)] == [True, True, True, False]
    assert lim.permitir("ip:2")[0] is True


def test_fila_com_teto():
    f = app.Fila(vagas=1, max_espera=2)
    assert f.entrar_na_espera() and f.entrar_na_espera()
    assert f.entrar_na_espera() is False
    f.sair_da_espera()
    assert f.entrar_na_espera() is True


# ---------------------------------------------------------------- tetos do processo
posix = pytest.mark.skipif(os.name != "posix", reason="vigia de memória usa /proc")


@posix
def test_teto_de_tempo_mata_o_processo(monkeypatch):
    monkeypatch.setattr(app, "TIMEOUT_CONVERSAO_S", 1)
    with pytest.raises(app.FalhaConversao) as e:
        app.executar_conversor("fiel", (FIX / "digital.pdf").read_bytes())
    assert e.value.motivo == "tempo"
    time.sleep(0.5)
    vivos = []
    for d in Path("/proc").iterdir():
        if d.name.isdigit():
            try:
                cmd = (d / "cmdline").read_bytes().replace(b"\0", b" ")
            except OSError:
                continue
            if b"tesseract" in cmd or b"conversor.py" in cmd:
                vivos.append(cmd)
    assert vivos == [], f"sobrou processo da conversão: {vivos}"


@posix
def test_teto_de_memoria_mata_o_processo(monkeypatch):
    monkeypatch.setattr(app, "LIMITE_MEMORIA_MB", 60)
    with pytest.raises(app.FalhaConversao) as e:
        app.executar_conversor("fiel", (FIX / "digital.pdf").read_bytes())
    assert e.value.motivo == "memoria"


def test_nao_deixa_pasta_temporaria(tmp_path, monkeypatch):
    monkeypatch.setattr(app.tempfile, "gettempdir", lambda: str(tmp_path))
    app.executar_conversor("editavel", (FIX / "digital.pdf").read_bytes())
    assert list(tmp_path.iterdir()) == []


def test_pdf_que_quebra_no_conversor_vira_falha_controlada():
    with pytest.raises(app.FalhaConversao) as e:
        app.executar_conversor("fiel", b"%PDF-1.7 lixo")
    assert e.value.motivo == "falha"
