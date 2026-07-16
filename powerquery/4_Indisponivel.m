let
    // ========= NOVA FONTE: arquivo do sistema oficial no SharePoint (xlsx ou csv) =========
    Pasta = "https://SEU_TENANT.sharepoint.com/sites/SEU_SITE/Documentos%20Compartilhados/CAMINHO/DA/PASTA/RELATORIOS/",
    Wide = fnParser(Web.Contents(Pasta & "INDISPONIVEL.xlsx"), "#(tab)"),
    ParseNum = (v) =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else let t = Text.Trim(Text.From(v))
             in if t = "" or t = "-" then null else Number.From(t, "pt-BR"),
    ColsMet = List.Select(Table.ColumnNames(Wide), each Text.StartsWith(_, "MET|")),
    Longo = Table.UnpivotOtherColumns(Wide,
        List.RemoveItems(Table.ColumnNames(Wide), ColsMet), "MET", "VALOR"),
    ComSaldo = Table.AddColumn(Longo, "SALDORCONTACONTBIL", each ParseNum([VALOR])),
    ComValor = Table.SelectRows(ComSaldo, each [SALDORCONTACONTBIL] <> null),
    ComConta = Table.AddColumn(ComValor, "ID_CONTA_CONTABIL", each Number.From(Text.Split([MET], "|"){1})),
    ComNoConta = Table.AddColumn(ComConta, "NO_CONTA_CONTABIL",
        each let s = Text.Split([MET], "|"){2}
             in if Text.StartsWith(s, "=") then s else "= " & s),
    Renome = Table.RenameColumns(ComNoConta, {
        {"Ação Governo", "ID_ACAO_PT"},
        {"Plano Orçamentário", "ID_PO"},
        {"Grupo Despesa", "ID_GRUPO_DESPESA_NADE"},
        {"Grupo Despesa (nome)", "NO_GRUPO_DESPESA_NADE"}}, MissingField.Ignore),
    // nomes de Ação/PO (necessários para os merges) vêm da execução completa
    Exec = fnParser(Web.Contents(Pasta & "EXECUCAO_COMPLETA.xlsx"), ","),
    MapaAcao = Table.Distinct(Table.SelectColumns(Exec, {"Ação Governo", "Ação Governo (nome)"})),
    MapaPO = Table.Distinct(Table.SelectColumns(Exec, {"Plano Orçamentário", "Plano Orçamentário (nome)"})),
    JAcao = Table.ExpandTableColumn(
        Table.NestedJoin(Renome, {"ID_ACAO_PT"}, MapaAcao, {"Ação Governo"}, "j1", JoinKind.LeftOuter),
        "j1", {"Ação Governo (nome)"}, {"NO_ACAO_PT"}),
    JPO = Table.ExpandTableColumn(
        Table.NestedJoin(JAcao, {"ID_PO"}, MapaPO, {"Plano Orçamentário"}, "j2", JoinKind.LeftOuter),
        "j2", {"Plano Orçamentário (nome)"}, {"NO_PO"}),
    ComNulos = List.Accumulate({"ID_UO", "ID_FUNCAO_PT", "ID_SUBFUNCAO_PT", "ID_PROGRAMA_PT"}, JPO,
        (t, c) => if List.Contains(Table.ColumnNames(t), c) then t
                  else Table.AddColumn(t, c, each null)),
    dbo_Indisponivel = Table.RemoveColumns(ComNulos, {"MET", "VALOR"}, MissingField.Ignore),
    // ========= DAQUI PARA BAIXO: PASSOS ORIGINAIS (inalterados) =========
    #"Tipo Alterado" = Table.TransformColumnTypes(dbo_Indisponivel,{{"ID_ACAO_PT", type any}, {"NO_ACAO_PT", type text}, {"ID_UO", Int64.Type}, {"ID_FUNCAO_PT", Int64.Type}, {"ID_SUBFUNCAO_PT", Int64.Type}, {"ID_PROGRAMA_PT", Int64.Type}, {"ID_PO", type any}, {"NO_PO", type text}, {"ID_GRUPO_DESPESA_NADE", Int64.Type}, {"NO_GRUPO_DESPESA_NADE", type text}, {"ID_CONTA_CONTABIL", Int64.Type}, {"NO_CONTA_CONTABIL", type text}, {"SALDORCONTACONTBIL", type number}}),
    #"Texto Extraído Após o Delimitador" = Table.TransformColumns(#"Tipo Alterado", {{"NO_CONTA_CONTABIL", each Text.AfterDelimiter(_, "= "), type text}}),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Texto Extraído Após o Delimitador",{{"ID_PO", type text}}),
    #"Consultas Mescladas" = Table.NestedJoin(#"Tipo Alterado1", {"NO_PO"}, PO, {"po"}, "PO", JoinKind.LeftOuter),
    #"PO Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas", "PO", {"id"}, {"PO.id"}),
    #"Consultas Mescladas1" = Table.NestedJoin(#"PO Expandido", {"NO_ACAO_PT"}, AÇ, {"aç"}, "AÇ", JoinKind.LeftOuter),
    #"AÇ Expandido" = Table.ExpandTableColumn(#"Consultas Mescladas1", "AÇ", {"id"}, {"AÇ.id"}),
    #"Tipo Alterado2" = Table.TransformColumnTypes(#"AÇ Expandido",{{"ID_PROGRAMA_PT", type text}}),
    #"Valor Substituído" = Table.ReplaceValue(#"Tipo Alterado2","32","0032",Replacer.ReplaceText,{"ID_PROGRAMA_PT"}),
    #"Valor Substituído1" = Table.ReplaceValue(#"Valor Substituído","910","0910",Replacer.ReplaceText,{"ID_PROGRAMA_PT"}),
    #"Tipo Alterado3" = Table.TransformColumnTypes(#"Valor Substituído1",{{"ID_ACAO_PT", type text}})
in
    #"Tipo Alterado3"