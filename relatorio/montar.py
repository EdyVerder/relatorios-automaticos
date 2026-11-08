"""Monta o e-mail (HTML com estilo em linha, que funciona no Outlook) e a planilha anexa."""
import datetime as dt
import html
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FAIXAS = [("Vencidos há mais de 30 dias", lambda d: d > 30, "#cf222e"),
          ("Vencidos de 8 a 30 dias", lambda d: 8 <= d <= 30, "#d1242f"),
          ("Vencidos até 7 dias", lambda d: 1 <= d <= 7, "#bc4c00"),
          ("Vencem hoje", lambda d: d == 0, "#9a6700"),
          ("Vencem nos próximos 7 dias", lambda d: d < 0, "#3b2a8f")]


def brl(v):
    return "R$ " + f"{v:,.2f}".replace(",", "_").replace(".", ",").replace("_", ".")


def agrupar(titulos):
    grupos = []
    for nome, regra, cor in FAIXAS:
        itens = [t for t in titulos if regra(t["dias_atraso"])]
        if itens:
            grupos.append({"nome": nome, "cor": cor, "itens": itens, "total": sum(t["valor"] for t in itens)})
    return grupos


def _situacao(dias):
    if dias > 0:
        return f"{dias} dia{'s' if dias > 1 else ''} em atraso"
    if dias == 0:
        return "vence hoje"
    return f"vence em {-dias} dia{'s' if dias < -1 else ''}"


def montar_html(grupos, hoje):
    e = html.escape
    blocos = []
    for g in grupos:
        linhas = "".join(
            f'<tr><td style="padding:8px 10px;border-bottom:1px solid #eaeef2;width:56%">{e(t["fornecedor"])}<br>'
            f'<span style="color:#59636e;font-size:12px">{e(t["documento"])} · {e(t["centro_custo"])}</span></td>'
            f'<td style="padding:8px 10px;border-bottom:1px solid #eaeef2;white-space:nowrap;width:24%">'
            f'{dt.date.fromisoformat(t["vencimento"]):%d/%m/%Y}<br><span style="color:{g["cor"]};font-size:12px">'
            f'{_situacao(t["dias_atraso"])}</span></td>'
            f'<td style="padding:8px 10px;border-bottom:1px solid #eaeef2;text-align:right;white-space:nowrap">'
            f'{brl(t["valor"])}</td></tr>' for t in g["itens"])
        blocos.append(
            f'<tr><td style="padding:22px 0 8px"><span style="display:inline-block;width:10px;height:10px;'
            f'border-radius:50%;background:{g["cor"]};margin-right:8px"></span><b style="font-size:15px">'
            f'{e(g["nome"])}</b> <span style="color:#59636e">· {len(g["itens"])} título'
            f'{"s" if len(g["itens"]) > 1 else ""} · {brl(g["total"])}</span></td></tr>'
            f'<tr><td><table role="presentation" width="100%" cellspacing="0" cellpadding="0" '
            f'style="border-collapse:collapse;font-size:14px">{linhas}</table></td></tr>')
    corpo = "".join(blocos) or ('<tr><td style="padding:24px 0;color:#1a7f37"><b>Nenhum título vencido ou a vencer '
                               'nos próximos 7 dias.</b></td></tr>')
    return f'''<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>Títulos a pagar</title></head>
<body style="margin:0;background:#f4f6fb;font-family:'Segoe UI',Helvetica,Arial,sans-serif;color:#1f2328">
<table role="presentation" width="100%" cellspacing="0" cellpadding="0"><tr><td align="center" style="padding:24px 12px">
<table role="presentation" width="680" cellspacing="0" cellpadding="0" style="max-width:680px;width:100%;background:#fff;
 border:1px solid #d1d9e0;border-radius:12px;overflow:hidden">
<tr><td style="background:#140c33;padding:22px 28px;border-bottom:3px solid #ff8a5b">
 <div style="color:#fff;font-size:20px;font-weight:600">Títulos a pagar · vencidos e a vencer</div>
 <div style="color:#c9c3ff;font-size:13px;margin-top:4px">Posição de {hoje:%d/%m/%Y} · planilha completa em anexo</div>
</td></tr>
<tr><td style="padding:4px 28px 24px"><table role="presentation" width="100%" cellspacing="0" cellpadding="0">{corpo}
</table></td></tr>
<tr><td style="padding:14px 28px;background:#f6f8fa;color:#59636e;font-size:12px">
 E-mail automático enviado pelo robô de relatórios. Dados fictícios.</td></tr>
</table></td></tr></table></body></html>'''


def montar_planilha(titulos, hoje):
    wb = Workbook()
    ws = wb.active
    ws.title = "Títulos"
    cab = ["Fornecedor", "Documento", "Centro de custo", "Vencimento", "Dias em atraso", "Valor"]
    ws.append(cab)
    for t in titulos:
        ws.append([t["fornecedor"], t["documento"], t["centro_custo"], dt.date.fromisoformat(t["vencimento"]),
                   max(t["dias_atraso"], 0), t["valor"]])
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="140C33")
        c.alignment = Alignment(vertical="center")
    for linha in ws.iter_rows(min_row=2):
        linha[3].number_format = "DD/MM/YYYY"
        linha[5].number_format = '"R$" #,##0.00'
    for i, largura in enumerate([34, 12, 30, 13, 15, 16], start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()
