"""Banco fictício de títulos a pagar e a consulta do relatório."""
import datetime as dt
import random
import sqlite3

CENTROS = ["Obra Demonstração 01", "Obra Demonstração 02", "Obra Demonstração 03", "Administrativo",
           "Comercial"]


def criar_banco(caminho, hoje, qtd=140, semente=3):
    rnd = random.Random(semente)
    con = sqlite3.connect(caminho)
    con.execute("DROP TABLE IF EXISTS titulos")
    con.execute("CREATE TABLE titulos (id INTEGER PRIMARY KEY, fornecedor TEXT, documento TEXT, centro_custo TEXT, "
                "vencimento TEXT, valor REAL, baixado INTEGER)")
    linhas = []
    for i in range(1, qtd + 1):
        venc = hoje + dt.timedelta(days=rnd.randint(-45, 20))
        baixado = 1 if venc < hoje and rnd.random() < 0.8 else 0
        linhas.append((i, f"Fornecedor Demonstração {rnd.randint(1, 30):03d}", f"NF {rnd.randint(1000, 99999)}",
                       rnd.choice(CENTROS), venc.isoformat(), round(rnd.uniform(300, 35000), 2), baixado))
    con.executemany("INSERT INTO titulos VALUES (?, ?, ?, ?, ?, ?, ?)", linhas)
    con.commit()
    return con


CONSULTA = """
SELECT fornecedor, documento, centro_custo, vencimento, valor,
       CAST(julianday(:hoje) - julianday(vencimento) AS INTEGER) AS dias_atraso
FROM titulos
WHERE baixado = 0 AND vencimento <= date(:hoje, '+7 day')
ORDER BY vencimento, fornecedor
"""


def titulos_em_foco(con, hoje):
    cur = con.execute(CONSULTA, {"hoje": hoje.isoformat()})
    nomes = [d[0] for d in cur.description]
    return [dict(zip(nomes, r)) for r in cur]
