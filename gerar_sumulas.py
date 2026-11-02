#!/usr/bin/env python3
"""
Gera súmulas de prova (uma página A4 por linha da planilha) para mostrar à câmera
e grampear no alvo.

Uso:
    pip install -r requirements.txt
    python gerar_sumulas.py                        -> usa provas.xlsx da mesma pasta
    python gerar_sumulas.py provas.xlsx            -> sumulas.pdf
    python gerar_sumulas.py provas.xlsx saida.pdf

Como .exe (gerado com gerar_exe.bat): dê duplo clique com provas.xlsx na mesma
pasta, ou arraste uma planilha para cima do GerarSumulas.exe.
"""
import sys
from datetime import date, datetime
from pathlib import Path

from openpyxl import load_workbook
from reportlab.lib.colors import HexColor, white
from reportlab.lib.enums import TA_LEFT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph

# ---------- Paleta (pouca tinta: só traços e texto coloridos, fundo branco) ----------
INK = HexColor("#111111")
GREY = HexColor("#6B6B6B")
LIGHT = HexColor("#B5B5B5")
RED = HexColor("#C8102E")      # destaque principal (etapa, regras, miss)
GREEN = HexColor("#1B7F3B")    # mosca
BLUE = HexColor("#1F4E9E")     # TRP
ORANGE = HexColor("#D2691E")   # humanóide

MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def formatar_etapa(v):
    """Data do Excel -> 'Nov/26'. Aceita também texto já pronto (ex.: 'Nov/26', '11/2026')."""
    if v is None or v == "":
        return ""
    if isinstance(v, (datetime, date)):
        return f"{MESES[v.month - 1]}/{v.year % 100:02d}"
    t = str(v).strip()
    partes = t.replace("-", "/").split("/")
    if len(partes) >= 2 and partes[-2].isdigit() and partes[-1].isdigit():
        m, a = int(partes[-2]), int(partes[-1])
        if 1 <= m <= 12:
            return f"{MESES[m - 1]}/{a % 100:02d}"
    return t[:1].upper() + t[1:]

W, H = A4
M = 34  # margem


# ---------- Leitura da planilha ----------
def ler_planilha(caminho):
    wb = load_workbook(caminho, data_only=True)
    ws = wb["Provas"] if "Provas" in wb.sheetnames else wb.active
    linhas = list(ws.iter_rows(values_only=True))
    cab = [str(c).strip() if c is not None else "" for c in linhas[0]]
    registros = []
    for row in linhas[1:]:
        d = {cab[i]: ("" if v is None else str(v).strip()) for i, v in enumerate(row) if i < len(cab)}
        if "Etapa" in cab:
            d["Etapa"] = formatar_etapa(row[cab.index("Etapa")])
        if not any(d.values()):
            continue
        if d.get("Imprimir", "S").upper().startswith("N"):
            continue
        registros.append(d)
    return registros


# ---------- Ícones vetoriais (só contorno) ----------
def icone_mira(c, x, y, r, cor):
    c.saveState(); c.setStrokeColor(cor); c.setLineWidth(1.6)
    c.circle(x, y, r); c.circle(x, y, r * 0.45)
    for dx, dy in ((1, 0), (-1, 0), (0, 1), (0, -1)):
        c.line(x + dx * r * 0.2, y + dy * r * 0.2, x + dx * r * 1.3, y + dy * r * 1.3)
    c.restoreState()


def icone_lapis(c, x, y, s, cor):
    c.saveState(); c.setStrokeColor(cor); c.setLineWidth(1.6); c.setLineJoin(1)
    p = c.beginPath()
    p.moveTo(x, y); p.lineTo(x + s * 0.25, y + s * 0.08); p.lineTo(x + s, y + s * 0.83)
    p.lineTo(x + s * 0.83, y + s); p.lineTo(x + s * 0.08, y + s * 0.25); p.close()
    c.drawPath(p, stroke=1, fill=0)
    c.line(x + s * 0.7, y + s * 0.53, x + s * 0.53, y + s * 0.7)
    c.restoreState()


def icone_humanoide(c, x, y, h, cor):
    """Silhueta de alvo humanoide (cabeça + tronco), base em (x,y)."""
    c.saveState(); c.setStrokeColor(cor); c.setLineWidth(1.8)
    w = h * 0.62
    c.circle(x, y + h * 0.82, h * 0.16)
    p = c.beginPath()
    p.moveTo(x - w / 2, y); p.lineTo(x - w / 2, y + h * 0.45)
    p.curveTo(x - w / 2, y + h * 0.62, x - w * 0.25, y + h * 0.64, x - w * 0.12, y + h * 0.64)
    p.lineTo(x + w * 0.12, y + h * 0.64)
    p.curveTo(x + w * 0.25, y + h * 0.64, x + w / 2, y + h * 0.62, x + w / 2, y + h * 0.45)
    p.lineTo(x + w / 2, y)
    c.drawPath(p, stroke=1, fill=0)
    c.restoreState()


def icone_trp(c, x, y, h, cor):
    """Alvo de tiro prático: placa retangular com zonas internas."""
    c.saveState(); c.setStrokeColor(cor); c.setLineWidth(1.8)
    w = h * 0.62
    c.roundRect(x - w / 2, y, w, h, 3)
    c.setLineWidth(1.1)
    c.roundRect(x - w * 0.3, y + h * 0.18, w * 0.6, h * 0.6, 2)
    c.roundRect(x - w * 0.13, y + h * 0.38, w * 0.26, h * 0.22, 1)
    c.restoreState()


def icone_4cores(c, x, y, h, cor=None):
    """Placa com 4 círculos (só contorno, cada um numa cor)."""
    c.saveState(); c.setStrokeColor(INK); c.setLineWidth(1.4)
    w = h * 0.8
    c.roundRect(x - w / 2, y, w, h, 3)
    r = h * 0.17
    cores = [RED, BLUE, GREEN, HexColor("#E0A800")]
    pos = [(-1, 1), (1, 1), (-1, -1), (1, -1)]
    c.setLineWidth(2)
    for (dx, dy), cc in zip(pos, cores):
        c.setStrokeColor(cc)
        c.circle(x + dx * w * 0.23, y + h / 2 + dy * h * 0.23, r)
    c.restoreState()


def icone_generico(c, x, y, h, cor):
    """Alvo genérico: círculos concêntricos com mira."""
    r = h * 0.45
    cx, cy = x, y + h / 2
    c.saveState(); c.setStrokeColor(cor); c.setLineWidth(1.6)
    c.circle(cx, cy, r); c.circle(cx, cy, r * 0.62); c.circle(cx, cy, r * 0.25)
    c.setLineWidth(1)
    c.line(cx - r * 1.15, cy, cx + r * 1.15, cy); c.line(cx, cy - r * 1.15, cx, cy + r * 1.15)
    c.restoreState()


def _normaliza(txt):
    import unicodedata
    t = unicodedata.normalize("NFKD", txt or "").encode("ascii", "ignore").decode().lower()
    return " ".join(t.split())


# Alvos conhecidos: (ícone, cor do nome, legenda pequena)
def tipo_alvo(nome):
    n = _normaliza(nome)
    if n.startswith("human"):
        return icone_humanoide, ORANGE, None
    if n == "trp" or n.startswith("trp ") or n.startswith("tiro rapido"):
        return icone_trp, BLUE, "Tiro Rápido Prático"
    if n in ("4 cores", "quatro cores", "4cores"):
        return icone_4cores, INK, None
    return icone_generico, GREY, None


# ---------- Helpers de texto ----------
def texto_ajustado(c, txt, x, y, larg_max, fonte, tam_max, tam_min=9, cor=INK):
    """Desenha uma linha reduzindo a fonte até caber."""
    t = tam_max
    while t > tam_min and c.stringWidth(txt, fonte, t) > larg_max:
        t -= 0.5
    while c.stringWidth(txt, fonte, t) > larg_max and len(txt) > 1:
        txt = txt[:-2] + "…"
    c.setFont(fonte, t); c.setFillColor(cor); c.drawString(x, y, txt)
    return t


def rotulo(c, txt, x, y, cor=GREY, tam=8.5):
    c.setFont("Helvetica-Bold", tam); c.setFillColor(cor)
    c.drawString(x, y, " ".join(txt.upper()) if len(txt) < 14 else txt.upper())


def campo(c, label, valor, x, y, larg, tam=15):
    rotulo(c, label, x, y)
    texto_ajustado(c, valor or "—", x, y - tam - 3, larg, "Helvetica-Bold", tam)


def paragrafo_ajustado(c, txt, x, y_top, larg, alt, tam_max=14, tam_min=8):
    tam = tam_max
    while True:
        st = ParagraphStyle("d", fontName="Helvetica", fontSize=tam, leading=tam * 1.32,
                            textColor=INK, alignment=TA_LEFT)
        p = Paragraph(txt.replace("\n", "<br/>"), st)
        _, h = p.wrap(larg, alt)
        if h <= alt or tam <= tam_min:
            p.drawOn(c, x, y_top - h)
            return
        tam -= 0.5


def com_unidade(valor, unidade):
    """'25' -> '25 m'; texto livre ('5 min', '7 jardas') é mantido como está."""
    v = (valor or "").strip()
    if not v:
        return "—"
    try:
        n = float(v.replace(",", "."))
        return f"{n:g} {unidade}".replace(".", ",")
    except ValueError:
        return v


# ---------- Página ----------
def desenhar_pagina(c, d, idx, total):
    larg = W - 2 * M
    y = H - M

    # Faixa superior: rótulo + página
    rotulo(c, "Registro de prova", M, y - 8, GREY, 9)
    c.setFont("Helvetica", 8.5); c.setFillColor(GREY)
    c.drawRightString(W - M, y - 8, f"{idx}/{total}")
    y -= 22

    # Prova (grande) + selo da Etapa
    selo_w, selo_h = 118, 74
    sx, sy = W - M - selo_w, y - selo_h
    c.setStrokeColor(RED); c.setLineWidth(3)
    c.roundRect(sx, sy, selo_w, selo_h, 8)
    c.setFillColor(RED); c.setFont("Helvetica-Bold", 10)
    c.drawCentredString(sx + selo_w / 2, sy + selo_h - 15, "E T A P A")
    etapa = d.get("Etapa", "") or "—"
    t = 34
    while t > 14 and c.stringWidth(etapa, "Helvetica-Bold", t) > selo_w - 14:
        t -= 1
    c.setFont("Helvetica-Bold", t); c.drawCentredString(sx + selo_w / 2, sy + 18, etapa)

    rotulo(c, "Prova", M, y - 8)
    texto_ajustado(c, d.get("Prova", ""), M, y - 44, larg - selo_w - 14, "Helvetica-Bold", 32, 14)
    c.setFillColor(GREY); c.setFont("Helvetica", 9)
    y = sy - 12

    # Régua vermelha
    c.setStrokeColor(RED); c.setLineWidth(3); c.line(M, y, W - M, y)
    y -= 24

    # Atirador (destaque para a câmera)
    rotulo(c, "Atirador", M, y)
    texto_ajustado(c, d.get("Atirador", ""), M, y - 32, larg, "Helvetica-Bold", 28, 14)
    y -= 58

    # Grade de dados
    col1 = larg * 0.62
    campo(c, "Clube de Tiro", d.get("Clube de Tiro", ""), M, y, col1 - 12)
    campo(c, "Cidade", d.get("Cidade", ""), M + col1, y, larg - col1)
    y -= 46
    c.setStrokeColor(LIGHT); c.setLineWidth(0.6); c.line(M, y + 10, W - M, y + 10)
    campo(c, "Arma", d.get("Arma", ""), M, y - 6, col1 - 12)

    # Alvo com ícone
    alvo = d.get("Alvo", "")
    icone, cor_alvo, legenda = tipo_alvo(alvo)
    ax = M + col1
    rotulo(c, "Alvo", ax, y - 6)
    ic_h = 30
    icone(c, ax + 12, y - 44, ic_h, cor_alvo)
    texto_ajustado(c, alvo or "—", ax + 34, y - 30, larg - col1 - 36, "Helvetica-Bold", 18, 10, cor_alvo)
    if legenda:
        c.setFont("Helvetica", 7.5); c.setFillColor(GREY)
        c.drawString(ax + 34, y - 42, legenda)
    y -= 62

    # Regras da prova: Distância / Tempo máx. / Tiros
    regras = [
        ("DISTÂNCIA", com_unidade(d.get("Distância", ""), "m")),
        ("TEMPO MÁX.", com_unidade(d.get("Tempo", ""), "s")),
        ("TIROS", d.get("Tiros", "") or "—"),
    ]
    rg_h, rg_gap = 58, 10
    rg_w = (larg - 2 * rg_gap) / 3
    for i, (lab, val) in enumerate(regras):
        rx = M + i * (rg_w + rg_gap)
        c.setStrokeColor(RED); c.setLineWidth(2)
        c.roundRect(rx, y - rg_h, rg_w, rg_h, 6)
        c.setFont("Helvetica-Bold", 9); c.setFillColor(RED)
        c.drawCentredString(rx + rg_w / 2, y - 15, " ".join(lab))
        t = 28
        while t > 12 and c.stringWidth(val, "Helvetica-Bold", t) > rg_w - 16:
            t -= 1
        c.setFont("Helvetica-Bold", t); c.setFillColor(INK)
        c.drawCentredString(rx + rg_w / 2, y - rg_h + 13, val)
    y -= rg_h + 16

    # Como atirar
    box_h = 128
    c.setStrokeColor(INK); c.setLineWidth(1.2)
    c.rect(M, y - box_h, larg, box_h)
    c.setStrokeColor(RED); c.setLineWidth(5)
    c.line(M + 2.5, y - box_h, M + 2.5, y)
    icone_mira(c, M + 24, y - 20, 8, RED)
    c.setFont("Helvetica-Bold", 13); c.setFillColor(RED)
    c.drawString(M + 42, y - 25, "COMO ATIRAR")
    paragrafo_ajustado(c, d.get("Descrição", "") or "—", M + 20, y - 40, larg - 36, box_h - 50)
    y -= box_h + 26

    # Resultado (preenchimento manual)
    icone_lapis(c, M, y - 4, 15, INK)
    c.setFont("Helvetica-Bold", 13); c.setFillColor(INK)
    c.drawString(M + 22, y, "RESULTADO")
    c.setFont("Helvetica-Oblique", 9); c.setFillColor(GREY)
    c.drawString(M + 112, y, "preencher à mão")
    y -= 12

    res_h = y - (M + 62)
    pts_w = larg * 0.40
    gap = 10
    # Pontos — caixa grande
    c.setStrokeColor(INK); c.setLineWidth(2.4)
    c.roundRect(M, y - res_h, pts_w, res_h, 6)
    c.setFont("Helvetica-Bold", 16); c.setFillColor(INK)
    c.drawString(M + 12, y - 24, "PONTOS")

    # Grade 2x2
    gx = M + pts_w + gap
    gw = (larg - pts_w - gap - gap) / 2
    gh = (res_h - gap) / 2
    caixas = [
        ("TEMPO", INK, "seg"), ("# DISPAROS", INK, None),
        ("# MOSCA", GREEN, None), ("# MISS", RED, None),
    ]
    for i, (lab, cor, unid) in enumerate(caixas):
        cx = gx + (i % 2) * (gw + gap)
        cy = y - (i // 2) * (gh + gap) - gh
        c.setStrokeColor(cor); c.setLineWidth(1.8)
        c.roundRect(cx, cy, gw, gh, 5)
        c.setFont("Helvetica-Bold", 12); c.setFillColor(cor)
        c.drawString(cx + 10, cy + gh - 20, lab)
        if lab == "# MOSCA":
            icone_mira(c, cx + gw - 18, cy + gh - 16, 5, GREEN)
        if lab == "# MISS":
            c.setStrokeColor(RED); c.setLineWidth(2)
            xx, yy = cx + gw - 18, cy + gh - 16
            c.line(xx - 6, yy - 6, xx + 6, yy + 6); c.line(xx - 6, yy + 6, xx + 6, yy - 6)
        if unid:
            c.setFont("Helvetica", 10); c.setFillColor(GREY)
            c.drawRightString(cx + gw - 10, cy + 10, unid)

    # Data / Assinatura
    by = M + 30
    c.setStrokeColor(GREY); c.setLineWidth(0.8)
    rotulo(c, "Data", M, by)
    c.line(M + 42, by - 2, M + 170, by - 2)
    c.setFont("Helvetica", 11); c.setFillColor(LIGHT)
    c.drawString(M + 70, by + 1, "/        /")
    rotulo(c, "Assinatura do atirador", M + 200, by)
    c.line(M + 200 + c.stringWidth("ASSINATURA DO ATIRADOR", "Helvetica-Bold", 8.5) + 8, by - 2, W - M, by - 2)

    c.setFont("Helvetica", 7); c.setFillColor(LIGHT)
    c.drawString(M, M, f"Gerado em {datetime.now():%d/%m/%Y %H:%M}")


def pasta_do_programa():
    """Pasta do .exe (quando empacotado) ou do script .py."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def gerar(origem, destino):
    if not origem.exists():
        raise SystemExit(f"Planilha não encontrada: {origem}\n"
                         f"Coloque 'provas.xlsx' na mesma pasta do programa, "
                         f"ou arraste a planilha para cima do programa.")
    regs = ler_planilha(origem)
    if not regs:
        raise SystemExit("Nenhuma linha para imprimir (verifique a coluna 'Imprimir').")
    try:
        c = canvas.Canvas(str(destino), pagesize=A4)
        c.setTitle("Súmulas de prova"); c.setAuthor("gerar_sumulas")
        for i, d in enumerate(regs, 1):
            c.setFillColor(white)
            desenhar_pagina(c, d, i, len(regs))
            c.showPage()
        c.save()
    except PermissionError:
        raise SystemExit(f"Não consegui gravar {destino.name}.\n"
                         f"Feche o PDF se ele estiver aberto e tente de novo.")
    print(f"{len(regs)} página(s) -> {destino}")
    return destino


def main():
    # Sem argumentos (duplo clique): usa provas.xlsx da pasta do programa.
    # Arrastar uma planilha sobre o .exe passa o caminho como argumento.
    origem = Path(sys.argv[1]) if len(sys.argv) > 1 else pasta_do_programa() / "provas.xlsx"
    destino = Path(sys.argv[2]) if len(sys.argv) > 2 else origem.with_name("sumulas.pdf")
    empacotado = getattr(sys, "frozen", False)
    try:
        pdf = gerar(origem, destino)
        if sys.platform == "win32":
            import os
            os.startfile(pdf)  # abre o PDF no leitor padrão
    except SystemExit as e:
        print(f"\nERRO: {e}")
        if empacotado:
            input("\nPressione Enter para fechar...")
        sys.exit(1)
    except Exception as e:  # erro inesperado: mostra antes de a janela fechar
        print(f"\nERRO inesperado: {e!r}")
        if empacotado:
            input("\nPressione Enter para fechar...")
        sys.exit(1)


if __name__ == "__main__":
    main()
