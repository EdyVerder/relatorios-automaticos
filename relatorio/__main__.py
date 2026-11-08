"""Relatório diário de títulos a pagar vencidos e a vencer, por e-mail, com a planilha completa em anexo.

Uso: python -m relatorio [--simular] [--hoje AAAA-MM-DD] [--config config.json]
Sem --simular, envia pelo SMTP das variáveis de ambiente SMTP_HOST, SMTP_PORTA, SMTP_USUARIO e SMTP_SENHA.
"""
import argparse
import datetime as dt
import json
import os
from pathlib import Path

from .dados import criar_banco, titulos_em_foco
from .enviar import enviar, mensagem
from .montar import agrupar, brl, montar_html, montar_planilha


def main():
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("--simular", action="store_true", help="grava o .eml e a prévia em vez de enviar")
    ap.add_argument("--hoje", type=dt.date.fromisoformat, default=dt.date.today())
    ap.add_argument("--config", default="config.json")
    ap.add_argument("--pasta", default="saida")
    args = ap.parse_args()

    cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
    pasta = Path(args.pasta)
    pasta.mkdir(parents=True, exist_ok=True)
    con = criar_banco(pasta / "financeiro_ficticio.db", args.hoje)   # na vida real: conexão com o ERP
    titulos = titulos_em_foco(con, args.hoje)
    grupos = agrupar(titulos)
    html = montar_html(grupos, args.hoje)
    anexo = montar_planilha(titulos, args.hoje)
    assunto = f"{cfg['assunto']} · {args.hoje:%d/%m/%Y}"
    msg = mensagem(assunto, html, cfg["para"], cfg.get("copia", []), anexo,
                   f"titulos_{args.hoje:%Y%m%d}.xlsx", os.environ.get("REMETENTE", cfg["remetente"]))

    for g in grupos:
        n = len(g["itens"])
        print(f"  {g['nome']:<28} {n:>3} título{'s' if n != 1 else ' '}  {brl(g['total']):>15}")
    if args.simular:
        (pasta / "previa.html").write_text(html, encoding="utf-8")
        (pasta / "relatorio.eml").write_bytes(bytes(msg))
        print(f"simulação: {pasta / 'relatorio.eml'} e {pasta / 'previa.html'} (nada foi enviado)")
    else:
        enviar(msg)
        print(f"enviado para {len(cfg['para'])} destinatário(s) e {len(cfg.get('copia', []))} em cópia")


if __name__ == "__main__":
    main()
