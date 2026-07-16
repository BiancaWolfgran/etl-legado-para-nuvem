# -*- coding: utf-8 -*-
"""
ETL DE EMERGÊNCIA — Painel AREA/COORDENACAO (Planejamento e Transferência de Recursos)
================================================================================
Substitui temporariamente o pipeline DaaS. Lê os 4 relatórios exportados do
sistema oficial de relatórios e produz:

  1. staging/Orcamento.csv       -> para carga em dbo.Orcamento      (SQL SEU_SERVIDOR_SQL_ANTIGO / SEU_BANCO)
  2. staging/Transferencia.csv   -> para carga em dbo.Transferencia
  3. staging/Indisponivel.csv    -> para carga em dbo.Indisponivel
  4. staging/emendas.xlsx        -> substitui o arquivo emendas.xlsx no SharePoint
                                    (sites/SEU_SITE/.../AREA/COORDENACAO/emendas.xlsx)

Uso:
  python etl_emergencia.py --entrada <pasta com os CSVs> --saida <pasta staging>

Para a carga no SQL, use o script carga_sql.py (gerado junto) ou o comando
BULK INSERT indicado no README.

Regras implementadas (engenharia reversa do Power Query do PBIX):
  - Cabeçalho duplo do sistema oficial (linha de atributos + linha de nomes de métrica)
  - Dimensões em pares código/descrição
  - Unpivot dos itens de informação (colunas 8,9,13,... -> linhas)
  - Meses sistema contábil federal 013/014 (apuração/encerramento) são PRESERVADOS
  - Linha "Total" do rodapé descartada
  - Números pt-BR ("1.234.567,89") convertidos
  - NO_CONTA_CONTABIL mantém o prefixo "= " (o Power Query remove com
    Text.AfterDelimiter, então o dado bruto precisa ter o prefixo)
  - NO_AUTOR_EMENDA mantém o sufixo " / EMENDA N" (o Power Query remove com
    Text.BeforeDelimiter)
"""

import argparse
import csv
import io
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

# ----------------------------------------------------------------------------
MESES = {"JAN": 1, "FEV": 2, "MAR": 3, "ABR": 4, "MAI": 5, "JUN": 6,
         "JUL": 7, "AGO": 8, "SET": 9, "OUT": 10, "NOV": 11, "DEZ": 12}

ANO_VIGENTE = datetime.now().year
INGESTION = datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def norm(s: str) -> str:
    """normaliza p/ comparação: sem acento, maiúsculo, trim"""
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode()
    return s.upper().strip()


def parse_num(v):
    """'1.234.567,89' -> 1234567.89 ; vazio -> None"""
    if v is None:
        return None
    v = str(v).strip().strip('"')
    if v == "" or v == "-":
        return None
    v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def parse_mes(v):
    """'JUL/2026' -> (2026, 7, 'JUL/2026'); '014/2016' -> (2016, 14, '014/2016')"""
    if not v:
        return None
    v = v.strip().strip('"')
    m = re.match(r"^([A-Z]{3}|\d{3})/(\d{4})$", norm(v))
    if not m:
        return None
    mes_raw, ano = m.group(1), int(m.group(2))
    mes = MESES.get(mes_raw, int(mes_raw) if mes_raw.isdigit() else None)
    if mes is None:
        return None
    return ano, mes, v


def ler_relatorio(path: Path):
    """
    Lê um export do sistema oficial de relatórios (csv ',' ou tsv), detecta o cabeçalho
    duplo e devolve (header1, header2, rows).
    """
    raw = path.read_bytes()
    for enc in ("latin-1", "utf-8-sig", "utf-8"):
        try:
            text = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    lines = text.splitlines()

    # detecta linha do cabeçalho: primeira linha que contém um atributo conhecido
    hdr_idx = None
    for i, l in enumerate(lines):
        ln = norm(l)
        if "MES LANCAMENTO" in ln or ln.startswith('"ACAO GOVERNO"'):
            hdr_idx = i
            break
    if hdr_idx is None:
        raise ValueError(f"{path.name}: cabeçalho não encontrado")

    # detecta delimitador na linha de cabeçalho
    delim = "\t" if lines[hdr_idx].count("\t") >= lines[hdr_idx].count(",") else ","

    body = "\n".join(lines[hdr_idx:])
    rows = list(csv.reader(io.StringIO(body), delimiter=delim))
    header1, header2 = rows[0], rows[1]
    data = [r for r in rows[2:] if any(c.strip() for c in r)]
    # descarta linha Total
    data = [r for r in data if norm(r[0]) != "TOTAL"]
    return header1, header2, data


def montar_colunas(header1, header2):
    """
    Constrói a lista de colunas finais.
    - dimensão com célula seguinte vazia no header1 => par (COD, NOME)
    - coluna numérica no header1 com nome no header2 => métrica (codigo, nome)
    Retorna lista de tuplas: ('dim', nome, 'cod'|'nome') ou ('met', codigo, nome)
    """
    cols = []
    n = len(header1)
    i = 0
    while i < n:
        h1 = (header1[i] or "").strip().strip('"')
        h2 = (header2[i] or "").strip().strip('"') if i < len(header2) else ""
        if h1 == "":
            i += 1
            continue
        if re.match(r"^\d+$", h1):                      # métrica pivotada
            cols.append(("met", i, h1, h2))
            i += 1
        else:                                           # dimensão
            # tem coluna de descrição? (próximo header1 vazio E não é métrica)
            if i + 1 < n and (header1[i + 1] or "").strip().strip('"') == "":
                cols.append(("dim2", i, h1))            # código + nome
                i += 2
            else:
                cols.append(("dim1", i, h1))            # só código
                i += 1
    return cols


def extrair(path: Path):
    """Devolve DataFrame 'wide' com dims nomeadas + métricas, e a lista de métricas."""
    h1, h2, data = ler_relatorio(path)
    cols = montar_colunas(h1, h2)
    registros = []
    metricas = [(c[2], c[3]) for c in cols if c[0] == "met"]
    for r in data:
        rec = {}
        for c in cols:
            if c[0] == "dim2":
                _, idx, nome = c
                rec[f"{nome}__cod"] = (r[idx] if idx < len(r) else "").strip().strip('"')
                rec[f"{nome}__nome"] = (r[idx + 1] if idx + 1 < len(r) else "").strip().strip('"')
            elif c[0] == "dim1":
                _, idx, nome = c
                rec[f"{nome}__cod"] = (r[idx] if idx < len(r) else "").strip().strip('"')
            else:
                _, idx, cod, nomemet = c
                rec[f"MET__{cod}"] = parse_num(r[idx] if idx < len(r) else None)
        registros.append(rec)
    return pd.DataFrame(registros), metricas


def unpivot(df, metricas, valor_col):
    """Wide -> long. Uma linha por (dims, item de informação) com valor não-nulo."""
    id_vars = [c for c in df.columns if not c.startswith("MET__")]
    out = []
    for cod, nome in metricas:
        col = f"MET__{cod}"
        if col not in df.columns:
            continue
        sub = df[df[col].notna()][id_vars].copy()
        sub["CO_ITEM_INFORMACAO"] = int(cod)
        sub["NO_ITEM_INFORMACAO"] = nome
        sub[valor_col] = df.loc[sub.index, col].values
        out.append(sub)
    return pd.concat(out, ignore_index=True) if out else pd.DataFrame()


def split_mes(df, col="Mês Lançamento__cod"):
    parsed = df[col].map(parse_mes)
    df["ID_ANO_LANC"] = parsed.map(lambda t: t[0] if t else None)
    df["ID_MES_LANC"] = parsed.map(lambda t: t[1] if t else None)
    df["SG_MES_COMPLETO"] = parsed.map(lambda t: t[2] if t else None)
    return df[df["ID_ANO_LANC"].notna()].copy()


# ============================================================================
def transf_orcamento(df_long):
    d = split_mes(df_long)
    out = pd.DataFrame({
        "ID_ANO_LANC":            d["ID_ANO_LANC"].astype("Int64"),
        "ID_MES_LANC":            d["ID_MES_LANC"].astype("Int64"),
        "SG_MES_COMPLETO":        d["SG_MES_COMPLETO"],
        "ID_PTRES":               d.get("PTRES__cod"),
        "ID_FUNCAO_PT":           None,
        "ID_SUBFUNCAO_PT":        None,
        "ID_PROGRAMA_PT":         d.get("Programa Governo__cod"),
        "NO_PROGRAMA_PT":         d.get("Programa Governo__nome"),
        "ID_ACAO_PT":             d.get("Ação Governo__cod"),
        "NO_ACAO_PT":             d.get("Ação Governo__nome"),
        "ID_LOCALIZADOR_GASTO_PT": None,
        "CO_PT":                  None,
        "ID_UG_EXEC":             d.get("UG Executora__cod"),
        "CO_UG":                  d.get("UG Executora__cod"),
        "NO_UG":                  d.get("UG Executora__nome"),
        "ID_UO":                  None,
        "ID_PO":                  d.get("Plano Orçamentário__cod"),
        "NO_PO":                  d.get("Plano Orçamentário__nome"),
        "ID_ORGAO_PI":            None,
        "ID_PI":                  d.get("PI__cod"),
        "NO_PI":                  None,
        "ID_GRUPO_DESPESA_NADE":  d.get("Grupo Despesa__cod"),
        "NO_GRUPO_DESPESA_NADE":  d.get("Grupo Despesa__nome"),
        "ID_IDUSO":               d.get("Iduso__cod"),
        "NO_IDUSO":               d.get("Iduso__nome"),
        "ID_IN_RESULTADO_LEI_CEOR": d.get("Resultado Lei__cod"),
        "NO_IN_RESULTADO_LEI_CEOR": d.get("Resultado Lei__nome"),
        "ID_ITEM_INFORMACAO":     d["CO_ITEM_INFORMACAO"].astype("Int64"),
        "CO_ITEM_INFORMACAO":     d["CO_ITEM_INFORMACAO"].astype("Int64"),
        "NO_ITEM_INFORMACAO":     d["NO_ITEM_INFORMACAO"],
        "SALDORITEMINFORMAO":     d["SALDORITEMINFORMAO"],
        "ID_UO0":                 None,
        "CO_UO":                  None,
        "NO_UO":                  None,
        "Nove_Sim_Nao":           None,
        "Classificacao_Unidade":  None,
        "Tipo_Instrumento":       None,
        "Filtro_Ano_Vigente":     d["ID_ANO_LANC"].map(lambda a: "Sim" if a == ANO_VIGENTE else "Não"),
        "ingestion_date":         INGESTION,
    })
    return out


def transf_transferencia(df_long):
    d = split_mes(df_long)
    out = pd.DataFrame({
        "ID_ANO_LANC":            d["ID_ANO_LANC"].astype("Int64"),
        "ID_MES_LANC":            d["ID_MES_LANC"].astype("Int64"),
        "SG_MES_COMPLETO":        d["SG_MES_COMPLETO"],
        "ID_PTRES":               d.get("PTRES__cod"),
        "ID_FUNCAO_PT":           None,
        "ID_SUBFUNCAO_PT":        None,
        "ID_PROGRAMA_PT":         d.get("Programa Governo__cod"),
        "NO_PROGRAMA_PT":         d.get("Programa Governo__nome"),
        "ID_ACAO_PT":             d.get("Ação Governo__cod"),
        "NO_ACAO_PT":             d.get("Ação Governo__nome"),
        "ID_LOCALIZADOR_GASTO_PT": None,
        "CO_PT":                  None,
        "ID_UG_EXEC":             d.get("UG Executora__cod"),
        "CO_UG":                  d.get("UG Executora__cod"),
        "NO_UG":                  d.get("UG Executora__nome"),
        "ID_UO":                  d.get("Unidade Orçamentária__cod"),
        "ID_UO0":                 d.get("Unidade Orçamentária__cod"),
        "CO_UO":                  d.get("Unidade Orçamentária__cod"),
        "NO_UO":                  d.get("Unidade Orçamentária__nome"),
        "ID_PO":                  d.get("Plano Orçamentário__cod"),
        "NO_PO":                  d.get("Plano Orçamentário__nome"),
        "ID_IDUSO":               d.get("Iduso__cod"),
        "NO_IDUSO":               d.get("Iduso__nome"),
        "ID_ORGAO_PI":            None,
        "ID_PI":                  d.get("PI__cod"),
        "NO_PI":                  None,
        "ID_GRUPO_DESPESA_NADE":  d.get("Grupo Despesa__cod"),
        "NO_GRUPO_DESPESA_NADE":  d.get("Grupo Despesa__nome"),
        "ID_ITEM_INFORMACAO":     d["CO_ITEM_INFORMACAO"].astype("Int64"),
        "CO_ITEM_INFORMACAO":     d["CO_ITEM_INFORMACAO"].astype("Int64"),
        "NO_ITEM_INFORMACAO":     d["NO_ITEM_INFORMACAO"],
        "SALDORITEMINFORMAO":     d["SALDORITEMINFORMAO"],
        "Nove_Sim_Nao":           None,
        "ID_PROGRAMA_PT0":        None, "ID_ACAO_PT0": None,
        "ID_FUNCAO_PT0":          None, "ID_SUBFUNCAO_PT0": None,
        "ID_PROGRAMA_PT1":        None, "ID_ACAO_PT1": None,
        "CO_PTRES":               d.get("PTRES__cod"),
        "ingestion_date":         INGESTION,
    })
    return out


def transf_indisponivel(df, metricas, mapa_acao, mapa_po):
    """Indisponível: métricas são CONTAS CONTÁBEIS pivotadas."""
    id_vars = [c for c in df.columns if not c.startswith("MET__")]
    out = []
    for cod, nome in metricas:
        col = f"MET__{cod}"
        if col not in df.columns:
            continue
        sub = df[df[col].notna()][id_vars].copy()
        sub["ID_CONTA_CONTABIL"] = int(cod)
        # Power Query remove o "= " com Text.AfterDelimiter -> manter prefixo
        sub["NO_CONTA_CONTABIL"] = nome if nome.startswith("=") else f"= {nome}"
        sub["SALDORCONTACONTBIL"] = df.loc[sub.index, col].values
        out.append(sub)
    d = pd.concat(out, ignore_index=True)
    acao = d.get("Ação Governo__cod", pd.Series([""] * len(d)))
    po_nome = d.get("Plano Orçamentário__nome")
    po_cod = d.get("Plano Orçamentário__cod", pd.Series([""] * len(d)))
    res = pd.DataFrame({
        "ID_ACAO_PT":            acao,
        # CSV traz só código; nome enriquecido a partir da EXECUÇÃO COMPLETA
        "NO_ACAO_PT":            acao.map(mapa_acao).fillna(""),
        "ID_UO":                 None,
        "ID_FUNCAO_PT":          None,
        "ID_SUBFUNCAO_PT":       None,
        "ID_PROGRAMA_PT":        None,
        "ID_PO":                 po_cod,
        "NO_PO":                 (po_nome if po_nome is not None else po_cod.map(mapa_po)).fillna("") if hasattr(po_nome if po_nome is not None else po_cod.map(mapa_po), 'fillna') else "",
        "ID_GRUPO_DESPESA_NADE": d.get("Grupo Despesa__cod"),
        "NO_GRUPO_DESPESA_NADE": d.get("Grupo Despesa__nome"),
        "ID_CONTA_CONTABIL":     d["ID_CONTA_CONTABIL"].astype("Int64"),
        "NO_CONTA_CONTABIL":     d["NO_CONTA_CONTABIL"],
        "SALDORCONTACONTBIL":    d["SALDORCONTACONTBIL"],
        "ingestion_date":        INGESTION,
    })
    return res


def transf_emendas(df_long):
    d = split_mes(df_long)
    autor = d.get("Autor Emendas Orçamento__cod", pd.Series([""] * len(d))).str.rstrip()
    out = pd.DataFrame({
        "ID_ANO_LANC":        d["ID_ANO_LANC"].astype("Int64"),
        "ID_MES_LANC":        d["ID_MES_LANC"].astype("Int64").astype(str),
        "SG_MES_COMPLETO":    d["SG_MES_COMPLETO"],
        "ID_UO":              d.get("Unidade Orçamentária__cod"),
        "CO_UO":              d.get("Unidade Orçamentária__cod"),
        "NO_UO":              d.get("Unidade Orçamentária__nome"),
        "ID_UG_EXEC":         d.get("UG Executora__cod"),
        "CO_UG":              d.get("UG Executora__cod"),
        "NO_UG":              d.get("UG Executora__nome"),
        "ID_RESULTADO_LEI":   d.get("Resultado Primário Lei__cod"),
        "NO_RESULTADO_LEI":   d.get("Resultado Primário Lei__nome"),
        "ID_AUTOR_EMENDA":    "",
        # Power Query faz Text.BeforeDelimiter(_, " /") -> manter " / EMENDA N"
        "NO_AUTOR_EMENDA":    autor,
        "ID_ACAO_PT":         d.get("Ação Governo__cod"),
        "NO_ACAO_PT":         "",
        "ID_UO0":             d.get("Unidade Orçamentária__cod"),
        "ID_FUNCAO_PT":       "",
        "ID_SUBFUNCAO_PT":    "",
        "ID_PROGRAMA_PT":     "",
        "ID_PO":              d.get("Plano Orçamentário__cod"),
        "NO_PO":              "",
        "ID_GRUPO_DESPESA_NADE": d.get("Grupo Despesa__cod"),
        "NO_GRUPO_DESPESA_NADE": d.get("Grupo Despesa__nome"),
        "ID_ITEM_INFORMACAO": d["CO_ITEM_INFORMACAO"].astype(str),
        "NO_ITEM_INFORMACAO": d["NO_ITEM_INFORMACAO"],
        "CO_ITEM_INFORMACAO": d["CO_ITEM_INFORMACAO"].astype(str),
        "SALDORCONTACONTBIL": d["SALDORCONTACONTBIL"],
        "Filtro_Ano_Vigente": d["ID_ANO_LANC"].map(lambda a: "Sim" if a == ANO_VIGENTE else "Não"),
    })
    return out


# ============================================================================
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--entrada", default=".", help="pasta com os 4 CSVs do sistema oficial")
    ap.add_argument("--saida", default="staging", help="pasta de saída")
    args = ap.parse_args()
    ent, sai = Path(args.entrada), Path(args.saida)
    sai.mkdir(parents=True, exist_ok=True)

    def acha(padrao):
        for f in ent.iterdir():
            if padrao in norm(f.name):
                return f
        raise FileNotFoundError(f"arquivo contendo '{padrao}' não encontrado em {ent}")

    f_exec = acha("EXECUCAO_COMPLETA")
    f_ted = acha("EXECUCAO_POR_TED")
    f_ind = acha("INDISPONIVEL")
    f_eme = acha("EMENDAS")

    # ---- EXECUÇÃO COMPLETA -> Orcamento
    df, mets = extrair(f_exec)
    mapa_acao = dict(zip(df.get("Ação Governo__cod", []), df.get("Ação Governo__nome", [])))
    mapa_po = dict(zip(df.get("Plano Orçamentário__cod", []), df.get("Plano Orçamentário__nome", [])))
    orc = transf_orcamento(unpivot(df, mets, "SALDORITEMINFORMAO"))
    orc.to_csv(sai / "Orcamento.csv", index=False, sep=";", encoding="utf-8-sig")
    print(f"Orcamento.csv      : {len(orc):>6} linhas  (de {len(df)} linhas wide, {len(mets)} itens de informação)")

    # ---- EXECUÇÃO POR TED -> Transferencia
    df, mets = extrair(f_ted)
    tra = transf_transferencia(unpivot(df, mets, "SALDORITEMINFORMAO"))
    tra.to_csv(sai / "Transferencia.csv", index=False, sep=";", encoding="utf-8-sig")
    print(f"Transferencia.csv  : {len(tra):>6} linhas  (de {len(df)} linhas wide, {len(mets)} itens de informação)")

    # ---- INDISPONÍVEL -> Indisponivel
    df, mets = extrair(f_ind)
    ind = transf_indisponivel(df, mets, mapa_acao, mapa_po)
    ind.to_csv(sai / "Indisponivel.csv", index=False, sep=";", encoding="utf-8-sig")
    print(f"Indisponivel.csv   : {len(ind):>6} linhas  (de {len(df)} linhas wide, {len(mets)} contas contábeis)")

    # ---- EMENDAS -> emendas.xlsx (SharePoint)
    df, mets = extrair(f_eme)
    eme = transf_emendas(unpivot(df, mets, "SALDORCONTACONTBIL"))
    with pd.ExcelWriter(sai / "emendas.xlsx", engine="openpyxl") as xw:
        eme.to_excel(xw, sheet_name="emendas", index=False)
    print(f"emendas.xlsx       : {len(eme):>6} linhas  (aba 'emendas')")

    print("\nOK. Arquivos em:", sai.resolve())


if __name__ == "__main__":
    main()
