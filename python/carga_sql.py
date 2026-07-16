# -*- coding: utf-8 -*-
"""
CARGA DE EMERGENCIA - SQL Server (SEU_SERVIDOR_SQL / SEU_BANCO)
====================================================================
Versao em pyodbc puro - sem pandas/SQLAlchemy.
Insere apenas as colunas que existem TANTO no CSV QUANTO na tabela;
colunas extras do CSV (ex.: as calculadas por DAX no PBIX) sao ignoradas
com aviso.

Estrategia segura (tudo numa transacao - qualquer erro desfaz tudo):
  1. Renomeia a tabela atual para <nome>_backup_AAAAMMDD
  2. Recria a tabela vazia com a MESMA estrutura da original
  3. Insere os dados do CSV
  Rollback manual, se precisar:
    DROP TABLE dbo.Orcamento;
    EXEC sp_rename 'dbo.Orcamento_backup_AAAAMMDD', 'Orcamento';

Requisitos:  pip install pyodbc
Uso:         python carga_sql.py --staging . --trusted
             python carga_sql.py --staging . --usuario SEU_USUARIO
"""

import argparse
import csv
import getpass
import sys
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path

import pyodbc

SERVIDOR = "SEU_SERVIDOR_SQL"
BANCO = "SEU_BANCO"
LOTE = 500  # linhas por executemany

TABELAS = {
    "Orcamento.csv": "Orcamento",
    "Transferencia.csv": "Transferencia",
    "Indisponivel.csv": "Indisponivel",
}


def conectar(args):
    candidatos = ["ODBC Driver 18 for SQL Server",
                  "ODBC Driver 17 for SQL Server",
                  "ODBC Driver 13 for SQL Server",
                  "SQL Server Native Client 11.0",
                  "SQL Server"]
    instalados = pyodbc.drivers()
    driver = next((d for d in candidatos if d in instalados), None)
    if driver is None:
        sys.exit("Nenhum driver ODBC de SQL Server encontrado. "
                 f"Instalados: {instalados}")
    print(f"Usando driver ODBC: {driver}")

    partes = [f"DRIVER={{{driver}}}", f"SERVER={SERVIDOR}", f"DATABASE={BANCO}"]
    if driver.startswith("ODBC Driver 18"):
        # o Driver 18 exige TLS por padrao e falha com certificado
        # autoassinado de servidor interno
        partes += ["Encrypt=no", "TrustServerCertificate=yes"]
    if args.trusted:
        partes.append("Trusted_Connection=yes")
    else:
        if not args.usuario:
            sys.exit("Informe --usuario SEU_USUARIO ou use --trusted.")
        senha = getpass.getpass(f"Senha de {args.usuario}@{SERVIDOR}: ")
        partes += [f"UID={args.usuario}", f"PWD={senha}"]

    cn = pyodbc.connect(";".join(partes), autocommit=False)
    fast = driver.startswith("ODBC Driver")  # o driver legado nao aguenta
    return cn, fast


def ler_csv(path: Path):
    """Le o CSV de staging ('' -> None). Retorna (colunas, linhas)."""
    with open(path, encoding="utf-8-sig", newline="") as fh:
        rd = csv.reader(fh, delimiter=";")
        cols = next(rd)
        rows = [[(c if c != "" else None) for c in r] for r in rd]
    return cols, rows


def colunas_tabela(cur, tabela):
    """Retorna lista [(nome, tipo, tamanho_max)] na ordem da tabela.
    tamanho_max: None para tipos nao-texto; -1 para varchar(max)."""
    cur.execute(
        "SELECT COLUMN_NAME, DATA_TYPE, CHARACTER_MAXIMUM_LENGTH "
        "FROM INFORMATION_SCHEMA.COLUMNS "
        "WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=? "
        "ORDER BY ORDINAL_POSITION", tabela)
    return [(r[0], r[1], r[2]) for r in cur.fetchall()]


def tabela_existe(cur, tabela):
    cur.execute(
        "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES "
        "WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME=?", tabela)
    return cur.fetchone()[0] > 0


def carregar_tabela(cur, tabela, csv_path, sufixo, fast, sem_backup):
    cols_csv, rows = ler_csv(csv_path)
    print(f"{tabela}: {len(rows)} linhas para carregar...")

    if not tabela_existe(cur, tabela):
        sys.exit(f"  ERRO: dbo.{tabela} nao existe no banco {BANCO}. "
                 "Confira se a migracao para a VM .125 trouxe as tabelas.")

    if sem_backup:
        cur.execute(f"TRUNCATE TABLE dbo.[{tabela}]")
    else:
        bkp = f"{tabela}_backup_{sufixo}"
        if tabela_existe(cur, bkp):
            cur.execute(f"DROP TABLE dbo.[{bkp}]")
        cur.execute(f"EXEC sp_rename 'dbo.{tabela}', '{bkp}'")
        cur.execute(f"SELECT * INTO dbo.[{tabela}] FROM dbo.[{bkp}] WHERE 1=0")
        print(f"  tabela atual preservada como dbo.{bkp}")

    info_sql = colunas_tabela(cur, tabela)
    cols_sql = [c[0] for c in info_sql]
    limites = {c[0]: c[2] for c in info_sql
               if c[2] is not None and c[2] != -1}  # so varchar(n)

    # interseccao: so insere colunas que existem dos dois lados
    comuns = [c for c in cols_csv if c in cols_sql]
    ignoradas_csv = [c for c in cols_csv if c not in cols_sql]
    nulas_sql = [c for c in cols_sql if c not in cols_csv]
    if ignoradas_csv:
        print(f"  aviso - colunas do CSV sem correspondente na tabela "
              f"(ignoradas): {ignoradas_csv}")
    if nulas_sql:
        print(f"  aviso - colunas da tabela ausentes no CSV "
              f"(ficarao NULL): {nulas_sql}")

    idx = [cols_csv.index(c) for c in comuns]
    dados = [[r[i] for i in idx] for r in rows]

    # converte cada coluna para o tipo Python compativel com o tipo SQL
    tipos = {c[0]: c[1] for c in info_sql}
    INTEIROS = {"int", "bigint", "smallint", "tinyint"}
    DECIMAIS = {"numeric", "decimal", "money", "smallmoney", "float", "real"}
    DATAS = {"datetime", "datetime2", "smalldatetime", "date"}

    def conv_int(v):
        return int(Decimal(v)) if v is not None else None

    def conv_dec(v):
        return Decimal(v) if v is not None else None

    def conv_dt(v):
        if v is None:
            return None
        for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d"):
            try:
                return datetime.strptime(v, fmt)
            except ValueError:
                continue
        raise ValueError(f"data invalida: {v!r}")

    for j, col in enumerate(comuns):
        t = tipos.get(col, "")
        if t in INTEIROS:
            fn, aviso_frac = conv_int, True
        elif t in DECIMAIS:
            fn, aviso_frac = conv_dec, False
        elif t in DATAS:
            fn, aviso_frac = conv_dt, False
        else:
            # texto: trunca se exceder varchar(n)
            lim = limites.get(col)
            if lim:
                estouros, exemplo = 0, None
                for r in dados:
                    v = r[j]
                    if v is not None and len(v) > lim:
                        exemplo = exemplo or v
                        r[j] = v[:lim]
                        estouros += 1
                if estouros:
                    print(f"  aviso - {col} (varchar({lim})): {estouros} "
                          f"valor(es) truncado(s). Ex.: '{exemplo[:60]}...'")
            continue

        fracoes = 0
        for r in dados:
            v = r[j]
            if v is None:
                continue
            if aviso_frac and "." in v and Decimal(v) != int(Decimal(v)):
                fracoes += 1
            try:
                r[j] = fn(v)
            except (ValueError, InvalidOperation) as e:
                sys.exit(f"  ERRO em {col} (tipo {t}): valor {v!r} "
                         f"nao conversivel ({e})")
        if fracoes:
            print(f"  aviso - {col} e '{t}' no banco mas ha {fracoes} "
                  f"valor(es) com casas decimais (parte fracionaria sera "
                  f"perdida). Considere: ALTER TABLE dbo.[{tabela}] "
                  f"ALTER COLUMN [{col}] numeric(20,2)")

    sql = (f"INSERT INTO dbo.[{tabela}] "
           f"({', '.join('[' + c + ']' for c in comuns)}) "
           f"VALUES ({', '.join('?' for _ in comuns)})")

    cur.fast_executemany = fast
    total = len(dados)
    for i in range(0, total, LOTE):
        cur.executemany(sql, dados[i:i + LOTE])
        print(f"  {min(i + LOTE, total)}/{total}", end="\r")
    print(f"\n  OK - dbo.{tabela} carregada com {total} linhas.")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--staging", default=".",
                    help="pasta com os CSVs gerados pelo etl_emergencia.py")
    ap.add_argument("--usuario", help="usuario SQL (omitir se usar --trusted)")
    ap.add_argument("--trusted", action="store_true",
                    help="autenticacao integrada do Windows")
    ap.add_argument("--sem-backup", action="store_true",
                    help="TRUNCATE direto, sem renomear backup")
    args = ap.parse_args()

    cn, fast = conectar(args)
    cur = cn.cursor()
    sufixo = datetime.now().strftime("%Y%m%d")
    try:
        for csv_nome, tabela in TABELAS.items():
            path = Path(args.staging) / csv_nome
            if not path.exists():
                print(f"[pulado] {csv_nome} nao encontrado em {args.staging}")
                continue
            carregar_tabela(cur, tabela, path, sufixo, fast, args.sem_backup)
        cn.commit()
        print("\nCarga concluida e confirmada (COMMIT).")
        print("Agora atualize o dataset no Power BI "
              "(lembrando de apontar a fonte para SEU_SERVIDOR_SQL).")
    except Exception:
        cn.rollback()
        print("\nERRO - transacao desfeita (ROLLBACK). "
              "Nada foi alterado no banco.")
        raise
    finally:
        cn.close()


if __name__ == "__main__":
    main()
