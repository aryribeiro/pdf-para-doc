"""Gera os PDFs de teste (sintéticos, sem dado pessoal).

- digital.pdf: 2 páginas A4 com texto de verdade (camada de texto) e um quadro.
- escaneado.pdf: 1 página A4 que é só uma imagem (simula scanner).
- senha.pdf: protegido por senha.
- pagina_gigante.pdf: 1 página de 5 x 5 m (bomba de pixels a 300 dpi).
- muitas_paginas.pdf: 120 páginas em branco com uma linha de texto.
- nao_e_pdf.pdf: um PNG renomeado.

Uso: python tests/fixtures/gerar_fixtures.py
"""
from pathlib import Path

import fitz

DIR = Path(__file__).resolve().parent

PARAGRAFOS = [
    "O cliente solicitou a revisão do contrato de prestação de serviços.",
    "A empresa respondeu em 5 dias úteis, conforme o prazo acordado.",
    "Em caso de dúvida, consulte o setor jurídico pelo canal oficial.",
    "1 Introdução",
    "Protocolo 2026/000123 registrado no sistema.",
]


def _pagina_texto(doc, titulo):
    page = doc.new_page(width=595, height=842)
    page.insert_text((72, 90), titulo, fontsize=20, fontname="helv")
    y = 140
    for p in PARAGRAFOS:
        page.insert_text((72, y), p, fontsize=12, fontname="helv")
        y += 28
    page.draw_rect(fitz.Rect(72, y + 10, 520, y + 90), color=(0, 0, 0), width=1)
    page.insert_text((84, y + 50), "Quadro de resumo: valor total R$ 1.250,00", fontsize=11)
    return page


def main():
    doc = fitz.open()
    _pagina_texto(doc, "Relatório de Atendimento")
    _pagina_texto(doc, "Anexo II")
    doc.save(DIR / "digital.pdf")

    # escaneado: rasteriza a página 1 e cola como imagem numa página nova
    pix = doc[0].get_pixmap(dpi=150)
    esc = fitz.open()
    pg = esc.new_page(width=595, height=842)
    pg.insert_image(pg.rect, stream=pix.tobytes("jpg", jpg_quality=80))
    esc.save(DIR / "escaneado.pdf")

    doc.save(DIR / "senha.pdf", encryption=fitz.PDF_ENCRYPT_AES_256,
             owner_pw="dono", user_pw="segredo")

    gig = fitz.open()
    g = gig.new_page(width=14173, height=14173)  # 5 m x 5 m em pontos
    g.insert_text((72, 90), "pagina gigante", fontsize=40)
    gig.save(DIR / "pagina_gigante.pdf")

    muitas = fitz.open()
    for i in range(120):
        m = muitas.new_page(width=595, height=842)
        m.insert_text((72, 90), f"Pagina {i + 1}", fontsize=12)
    muitas.save(DIR / "muitas_paginas.pdf")

    (DIR / "nao_e_pdf.pdf").write_bytes(pix.tobytes("png"))
    doc.close()
    esc.close()
    print("fixtures geradas em", DIR)


if __name__ == "__main__":
    main()
