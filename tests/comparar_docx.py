"""Compara dois conjuntos de DOCX parte a parte (ignora docProps/core.xml).

Uso: python tests/comparar_docx.py <pasta_a> <pasta_b>
"""
import sys
import zipfile
from pathlib import Path

a_dir, b_dir = Path(sys.argv[1]), Path(sys.argv[2])
falhou = False
for a in sorted(a_dir.glob("*.docx")):
    b = b_dir / a.name
    if not b.is_file():
        continue
    za, zb = zipfile.ZipFile(a), zipfile.ZipFile(b)
    nomes = set(za.namelist()) | set(zb.namelist())
    dif = [n for n in sorted(nomes) if n != "docProps/core.xml"
           and (n not in za.namelist() or n not in zb.namelist() or za.read(n) != zb.read(n))]
    print(a.stem, "IGUAL" if not dif else f"DIFERENTE: {dif}")
    falhou |= bool(dif)
sys.exit(1 if falhou else 0)
