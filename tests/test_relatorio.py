import datetime as dt
import email
from io import BytesIO

from openpyxl import load_workbook

from relatorio.dados import criar_banco, titulos_em_foco
from relatorio.enviar import mensagem
from relatorio.montar import agrupar, brl, montar_html, montar_planilha

HOJE = dt.date(2026, 11, 7)


def titulo(dias, valor=100.0):
    return {"fornecedor": "Fornecedor Demonstração 001", "documento": "NF 1", "centro_custo": "Administrativo",
            "vencimento": (HOJE - dt.timedelta(days=dias)).isoformat(), "valor": valor, "dias_atraso": dias}


def test_faixas():
    grupos = agrupar([titulo(40), titulo(10), titulo(3), titulo(0), titulo(-2), titulo(-5, 50)])
    assert [(g["nome"], len(g["itens"])) for g in grupos] == [
        ("Vencidos há mais de 30 dias", 1), ("Vencidos de 8 a 30 dias", 1), ("Vencidos até 7 dias", 1),
        ("Vencem hoje", 1), ("Vencem nos próximos 7 dias", 2)]
    assert grupos[-1]["total"] == 150


def test_brl():
    assert brl(1234.5) == "R$ 1.234,50"


def test_consulta_so_traz_em_aberto_ate_7_dias(tmp_path):
    con = criar_banco(tmp_path / "f.db", HOJE)
    titulos = titulos_em_foco(con, HOJE)
    assert titulos and all(t["dias_atraso"] >= -7 for t in titulos)
    abertos = con.execute("SELECT COUNT(*) FROM titulos WHERE baixado = 0 AND vencimento <= date(?, '+7 day')",
                          (HOJE.isoformat(),)).fetchone()[0]
    assert len(titulos) == abertos


def test_html_escapa_e_mostra_grupos():
    t = titulo(3)
    t["fornecedor"] = "Fornecedor <Teste> & Cia"
    html = montar_html(agrupar([t]), HOJE)
    assert "Fornecedor &lt;Teste&gt; &amp; Cia" in html and "Vencidos até 7 dias" in html


def test_html_sem_titulos():
    assert "Nenhum título vencido" in montar_html([], HOJE)


def test_planilha_e_mensagem():
    titulos = [titulo(12, 999.9), titulo(-1)]
    xlsx = montar_planilha(titulos, HOJE)
    ws = load_workbook(BytesIO(xlsx)).active
    assert ws.max_row == 3 and ws["F2"].value == 999.9 and ws["E3"].value == 0
    msg = mensagem("Assunto", "<p>oi</p>", ["a@exemplo.com.br"], ["b@exemplo.com.br"], xlsx, "t.xlsx",
                   "r@exemplo.com.br")
    lida = email.message_from_bytes(bytes(msg))
    anexos = [p.get_filename() for p in lida.walk() if p.get_filename()]
    assert anexos == ["t.xlsx"] and lida["Cc"] == "b@exemplo.com.br"
