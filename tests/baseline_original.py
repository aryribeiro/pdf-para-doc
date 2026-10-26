"""Roda o aplicativo.py ORIGINAL (checkpoint-pre-devsecops) sobre um PDF e
grava o DOCX + o pico de memória. Cada chamada num processo novo, para o
pico medido ser só desta conversão.

Uso: python tests/baseline_original.py <app.py> <modo:fiel|editavel> <pdf> <saida.docx>
"""
import importlib.util
import io
import json
import resource
import sys
import time

app_path, modo, pdf, saida = sys.argv[1:5]
spec = importlib.util.spec_from_file_location("app_original", app_path)
app = importlib.util.module_from_spec(spec)
spec.loader.exec_module(app)

t0 = time.time()
with open(pdf, "rb") as f:
    buf = io.BytesIO(f.read())
fn = app.modo_copia_fiel if modo == "fiel" else app.modo_texto_editavel
docx = fn(buf)
with open(saida, "wb") as f:
    f.write(docx)
print(json.dumps({
    "segundos": round(time.time() - t0, 1),
    "pico_python_mb": resource.getrusage(resource.RUSAGE_SELF).ru_maxrss // 1024,
    "pico_tesseract_mb": resource.getrusage(resource.RUSAGE_CHILDREN).ru_maxrss // 1024,
    "bytes": len(docx),
}))
