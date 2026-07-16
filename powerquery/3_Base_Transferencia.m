let
    // ========= NOVA FONTE: arquivo do sistema oficial no SharePoint (xlsx ou csv) =========
    Pasta = "https://SEU_TENANT.sharepoint.com/sites/SEU_SITE/Documentos%20Compartilhados/CAMINHO/DA/PASTA/RELATORIOS/",
    Wide = fnParser(Web.Contents(Pasta & "EXECUCAO_TED.xlsx"), ","),
    ParseNum = (v) =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else let t = Text.Trim(Text.From(v))
             in if t = "" or t = "-" then null else Number.From(t, "pt-BR"),
    ColsMet = List.Select(Table.ColumnNames(Wide), each Text.StartsWith(_, "MET|")),
    Longo = Table.UnpivotOtherColumns(Wide,
        List.RemoveItems(Table.ColumnNames(Wide), ColsMet), "MET", "VALOR"),
    ComSaldo = Table.AddColumn(Longo, "SALDORITEMINFORMAO", each ParseNum([VALOR])),
    ComValor = Table.SelectRows(ComSaldo, each [SALDORITEMINFORMAO] <> null),
    ComCO = Table.AddColumn(ComValor, "CO_ITEM_INFORMACAO", each Number.From(Text.Split([MET], "|"){1})),
    ComNO = Table.AddColumn(ComCO, "NO_ITEM_INFORMACAO", each Text.Split([MET], "|"){2}),
    ComIDitem = Table.AddColumn(ComNO, "ID_ITEM_INFORMACAO", each [CO_ITEM_INFORMACAO]),
    MapaMes = [JAN=1, FEV=2, MAR=3, ABR=4, MAI=5, JUN=6, JUL=7, AGO=8, SET=9, OUT=10, NOV=11, DEZ=12],
    MesTxt = Table.TransformColumns(ComIDitem, {{"Mês Lançamento", each Text.Trim(Text.From(_)), type text}}),
    ComAno = Table.AddColumn(MesTxt, "ID_ANO_LANC", each Number.From(Text.AfterDelimiter([Mês Lançamento], "/"))),
    ComMes = Table.AddColumn(ComAno, "ID_MES_LANC",
        each let p = Text.Trim(Text.BeforeDelimiter([Mês Lançamento], "/"))
             in try Number.From(p) otherwise Record.Field(MapaMes, Text.Upper(p))),
    Renome = Table.RenameColumns(ComMes, {
        {"Mês Lançamento", "SG_MES_COMPLETO"},
        {"UG Executora", "CO_UG"},
        {"UG Executora (nome)", "NO_UG"},
        {"Programa Governo", "ID_PROGRAMA_PT"},
        {"Programa Governo (nome)", "NO_PROGRAMA_PT"},
        {"Ação Governo", "ID_ACAO_PT"},
        {"Ação Governo (nome)", "NO_ACAO_PT"},
        {"Plano Orçamentário", "ID_PO"},
        {"Plano Orçamentário (nome)", "NO_PO"},
        {"PTRES", "ID_PTRES"},
        {"Resultado Lei", "ID_IN_RESULTADO_LEI_CEOR"},
        {"Resultado Lei (nome)", "NO_IN_RESULTADO_LEI_CEOR"},
        {"Iduso", "ID_IDUSO"},
        {"Iduso (nome)", "NO_IDUSO"},
        {"Grupo Despesa", "ID_GRUPO_DESPESA_NADE"},
        {"Grupo Despesa (nome)", "NO_GRUPO_DESPESA_NADE"},
        {"PI", "ID_PI"},
        {"Unidade Orçamentária", "CO_UO"},
        {"Unidade Orçamentária (nome)", "NO_UO"}}, MissingField.Ignore),
    ComExec = Table.AddColumn(Renome, "ID_UG_EXEC", each [CO_UG]),
    ComUO = Table.AddColumn(ComExec, "ID_UO", each [CO_UO]),
    PassoExtra = Table.AddColumn(ComUO, "ID_UO0", each [CO_UO]),
    ComNulos = List.Accumulate({"ID_FUNCAO_PT", "ID_SUBFUNCAO_PT", "ID_LOCALIZADOR_GASTO_PT", "CO_PT", "ID_ORGAO_PI", "NO_PI"}, PassoExtra,
        (t, c) => if List.Contains(Table.ColumnNames(t), c) then t
                  else Table.AddColumn(t, c, each null)),
    dbo_Transferencia = Table.RemoveColumns(ComNulos,
        {"MET", "VALOR"} & List.Select(Table.ColumnNames(ComNulos),
            each Text.StartsWith(_, "Autor Emendas")), MissingField.Ignore),
    // ========= DAQUI PARA BAIXO: PASSOS ORIGINAIS (inalterados) =========
    #"Colunas Renomeadas" = Table.RenameColumns(dbo_Transferencia,{{"ID_PROGRAMA_PT", "Cod Programa Orçamentário"}, {"NO_PROGRAMA_PT", "Título do Programa Orçamentário"}}),
    #"Linhas Filtradas" = Table.SelectRows(#"Colunas Renomeadas", each true),
    #"Tipo Alterado1" = Table.TransformColumnTypes(#"Linhas Filtradas",{{"SALDORITEMINFORMAO", Currency.Type}}),
    #"Erros Substituídos" = Table.ReplaceErrorValues(#"Tipo Alterado1", {{"SALDORITEMINFORMAO", 0}}),
    #"Dividir Coluna por Delimitador" = Table.SplitColumn(#"Erros Substituídos", "SG_MES_COMPLETO", Splitter.SplitTextByEachDelimiter({"/"}, QuoteStyle.Csv, true), {"SG_MES_COMPLETO.1", "SG_MES_COMPLETO.2"}),
    #"Colunas Renomeadas1" = Table.RenameColumns(#"Dividir Coluna por Delimitador",{{"SG_MES_COMPLETO.2", "ano"}}),
    #"Tipo Alterado2" = Table.TransformColumnTypes(#"Colunas Renomeadas1",{{"ano", Int64.Type}, {"SG_MES_COMPLETO.1", type text}}),
    #"Colunas Removidas" = Table.RemoveColumns(#"Tipo Alterado2",{"SG_MES_COMPLETO.1", "ano"}),
    #"Texto Aparado" = Table.TransformColumns(Table.TransformColumnTypes(#"Colunas Removidas", {{"ID_ANO_LANC", type text}}, "pt-BR"),{{"ID_ANO_LANC", Text.Trim, type text}}),
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Texto Aparado",{{"ID_ANO_LANC", Int64.Type}, {"ID_MES_LANC", Int64.Type}, {"ID_PTRES", type text}, {"ID_FUNCAO_PT", Int64.Type}, {"ID_SUBFUNCAO_PT", Int64.Type}, {"Cod Programa Orçamentário", Int64.Type}, {"ID_ACAO_PT", type text}, {"ID_LOCALIZADOR_GASTO_PT", Int64.Type}, {"CO_PT", type text}, {"ID_UG_EXEC", Int64.Type}, {"CO_UG", Int64.Type}, {"NO_UG", type text}, {"Título do Programa Orçamentário", type text}, {"NO_ACAO_PT", type text}, {"ID_UO", Int64.Type}, {"ID_PO", type text}, {"NO_PO", type text}, {"ID_IDUSO", Int64.Type}, {"NO_IDUSO", type text}, {"ID_ORGAO_PI", Int64.Type}, {"ID_PI", type text}, {"NO_PI", type text}, {"ID_GRUPO_DESPESA_NADE", Int64.Type}, {"NO_GRUPO_DESPESA_NADE", type text}, {"ID_UO0", Int64.Type}, {"CO_UO", Int64.Type}, {"NO_UO", type text}, {"ID_ITEM_INFORMACAO", Int64.Type}, {"NO_ITEM_INFORMACAO", type text}, {"CO_ITEM_INFORMACAO", Int64.Type}}),
    #"Coluna Mesclada Inserida" = Table.AddColumn(#"Tipo Alterado", "chave_ptres", each Text.Combine({Text.From([ID_ANO_LANC], "pt-BR"), [ID_PTRES]}, ""), type text)
in
    #"Coluna Mesclada Inserida"