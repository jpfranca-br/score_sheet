#!/usr/bin/env python3
"""
Versão web do gerador de súmulas: a pessoa envia o Excel e recebe o PDF.

Reaproveita o layout de gerar_sumulas.py (que fica inalterado, na mesma pasta).

Desenvolvimento:
    pip install -r requirements-web.txt
    python app_web.py                      -> http://127.0.0.1:8000

Produção (servidor): ver DEPLOY.md (gunicorn + serviço do sistema).
"""
import io
import re
from datetime import datetime
from pathlib import Path

from flask import Flask, Response, render_template_string, request, send_file
from reportlab.lib.colors import white
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import gerar_sumulas as gs

PASTA = Path(__file__).resolve().parent
MODELO = PASTA / "provas.xlsx"
MAX_MB = 5          # tamanho máximo do upload
MAX_PAGINAS = 300   # limite de linhas/páginas por arquivo

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_MB * 1024 * 1024

PAGINA = """<!doctype html>
<html lang="pt-BR">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Gerador de Súmulas</title>
<style>
  :root { --ink:#111; --grey:#6b6b6b; --line:#d0d0d0; --red:#C8102E; --bg:#fff; }
  * { box-sizing:border-box; }
  body { margin:0; background:var(--bg); color:var(--ink);
         font:16px/1.5 system-ui, -apple-system, "Segoe UI", Roboto, Arial, sans-serif; }
  main { max-width:620px; margin:0 auto; padding:40px 16px 60px; }
  .kicker { font-size:12px; letter-spacing:.2em; color:var(--grey); font-weight:700; }
  h1 { font-size:32px; margin:4px 0 6px; }
  .rule { height:3px; background:var(--red); margin:16px 0 28px; }
  p { color:#333; }
  form { margin-top:20px; }
  .drop { display:block; border:2px dashed var(--line); border-radius:10px; padding:34px 20px;
          text-align:center; cursor:pointer; transition:border-color .15s, background .15s; }
  .drop:hover, .drop.over { border-color:var(--red); background:#fff7f8; }
  .drop strong { display:block; font-size:18px; }
  .drop span { color:var(--grey); font-size:14px; }
  .drop input { display:none; }
  #nome { margin-top:10px; font-weight:600; min-height:1.5em; }
  button { margin-top:16px; width:100%; padding:14px; font-size:17px; font-weight:700;
           color:#fff; background:var(--red); border:0; border-radius:8px; cursor:pointer; }
  button:disabled { background:#bbb; cursor:not-allowed; }
  .erro { border:2px solid var(--red); border-radius:8px; padding:12px 14px; margin:20px 0 0;
          color:var(--red); font-weight:600; }
  .modelo { margin-top:28px; padding-top:18px; border-top:1px solid var(--line); font-size:15px; }
  a { color:var(--red); font-weight:600; }
  small { color:var(--grey); }
</style>
</head>
<body>
<main>
  <div class="kicker">REGISTRO DE PROVA</div>
  <h1>Gerador de Súmulas</h1>
  <div class="rule"></div>
  <p>Envie a planilha preenchida e receba o PDF pronto para imprimir: uma página por linha.</p>

  {% if erro %}<div class="erro">{{ erro }}</div>{% endif %}

  <form method="post" action="{{ url_for('gerar') }}" enctype="multipart/form-data">
    <label class="drop" id="drop">
      <strong>Clique para escolher o Excel</strong>
      <span>ou arraste o arquivo .xlsx para cá (até {{ max_mb }} MB)</span>
      <input type="file" name="planilha" id="arquivo" accept=".xlsx" required>
      <div id="nome"></div>
    </label>
    <button type="submit" id="btn" disabled>Gerar PDF</button>
  </form>

  <div class="modelo">
    Primeira vez? <a href="{{ url_for('modelo') }}">Baixe a planilha modelo</a>
    e preencha a aba <b>Provas</b>.<br>
    <small>A planilha é processada na hora e não fica guardada no servidor.</small>
  </div>
</main>
<script>
  const drop = document.getElementById('drop'), inp = document.getElementById('arquivo'),
        nome = document.getElementById('nome'), btn = document.getElementById('btn');
  const mostrar = () => { const f = inp.files[0]; nome.textContent = f ? f.name : ''; btn.disabled = !f; };
  inp.addEventListener('change', mostrar);
  ['dragenter','dragover'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.add('over'); }));
  ['dragleave','drop'].forEach(e => drop.addEventListener(e, ev => { ev.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', ev => { if (ev.dataTransfer.files.length) { inp.files = ev.dataTransfer.files; mostrar(); } });
  document.querySelector('form').addEventListener('submit', () => {
    btn.textContent = 'Gerando…';
    setTimeout(() => { btn.textContent = 'Gerar PDF'; }, 4000);  // o download não recarrega a página
  });
</script>
</body>
</html>"""


def pagina(erro=None, status=200):
    html = render_template_string(PAGINA, erro=erro, max_mb=MAX_MB)
    return Response(html, status=status, mimetype="text/html")


def montar_pdf(arquivo):
    """Lê o Excel (arquivo em memória) e devolve (bytes do PDF, nº de páginas)."""
    regs = gs.ler_planilha(arquivo)
    if not regs:
        raise ValueError("Nenhuma linha para imprimir. Verifique a aba 'Provas' e a coluna 'Imprimir'.")
    if len(regs) > MAX_PAGINAS:
        raise ValueError(f"A planilha tem {len(regs)} linhas; o limite é {MAX_PAGINAS} por arquivo.")
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    c.setTitle("Súmulas de prova"); c.setAuthor("Gerador de Súmulas")
    for i, d in enumerate(regs, 1):
        c.setFillColor(white)
        gs.desenhar_pagina(c, d, i, len(regs))
        c.showPage()
    c.save()
    return buf.getvalue(), len(regs)


@app.get("/")
def inicio():
    return pagina()


@app.post("/gerar")
def gerar():
    f = request.files.get("planilha")
    if not f or not f.filename:
        return pagina("Escolha um arquivo .xlsx.", 400)
    if not f.filename.lower().endswith(".xlsx"):
        return pagina("O arquivo precisa ser um Excel .xlsx.", 400)
    dados = f.read()
    if dados[:2] != b"PK":  # todo .xlsx é um arquivo ZIP
        return pagina("O arquivo não parece ser um Excel .xlsx válido.", 400)
    try:
        pdf, n = montar_pdf(io.BytesIO(dados))
    except ValueError as e:
        return pagina(str(e), 400)
    except Exception:
        app.logger.exception("Falha ao processar planilha")
        return pagina("Não consegui ler esta planilha. Confira se ela segue o modelo "
                      "(aba 'Provas' e cabeçalhos na linha 1).", 400)

    base = re.sub(r"[^A-Za-z0-9_-]+", "_", Path(f.filename).stem).strip("_") or "provas"
    nome = f"sumulas_{base}_{datetime.now():%Y%m%d_%H%M}.pdf"
    app.logger.info("PDF gerado: %d página(s)", n)
    return send_file(io.BytesIO(pdf), mimetype="application/pdf",
                     as_attachment=True, download_name=nome)


@app.get("/modelo")
def modelo():
    if not MODELO.exists():
        return pagina("Planilha modelo indisponível no servidor.", 404)
    return send_file(MODELO, as_attachment=True, download_name="provas_modelo.xlsx")


@app.get("/saude")
def saude():
    return {"status": "ok"}


@app.errorhandler(413)
def grande_demais(_):
    return pagina(f"Arquivo maior que {MAX_MB} MB.", 413)


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8000, debug=False)
