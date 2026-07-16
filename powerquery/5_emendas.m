let
    // ========= NOVA FONTE: arquivo do sistema oficial no SharePoint (xlsx ou csv) =========
    Pasta = "https://SEU_TENANT.sharepoint.com/sites/SEU_SITE/Documentos%20Compartilhados/CAMINHO/DA/PASTA/RELATORIOS/",
    Wide = fnParser(Web.Contents(Pasta & "EMENDAS.xlsx"), "#(tab)"),
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
    ComCO = Table.AddColumn(ComValor, "CO_ITEM_INFORMACAO", each Text.Split([MET], "|"){1}),
    ComNO = Table.AddColumn(ComCO, "NO_ITEM_INFORMACAO", each Text.Split([MET], "|"){2}),
    ComIDitem = Table.AddColumn(ComNO, "ID_ITEM_INFORMACAO", each [CO_ITEM_INFORMACAO]),
    MapaMes = [JAN=1, FEV=2, MAR=3, ABR=4, MAI=5, JUN=6, JUL=7, AGO=8, SET=9, OUT=10, NOV=11, DEZ=12],
    MesTxt = Table.TransformColumns(ComIDitem, {{"Mês Lançamento", each Text.Trim(Text.From(_)), type text}}),
    ComAno = Table.AddColumn(MesTxt, "ID_ANO_LANC", each Number.From(Text.AfterDelimiter([Mês Lançamento], "/"))),
    ComMes = Table.AddColumn(ComAno, "ID_MES_LANC",
        each let p = Text.Trim(Text.BeforeDelimiter([Mês Lançamento], "/"))
             in Text.From(try Number.From(p) otherwise Record.Field(MapaMes, Text.Upper(p)))),
    Renome = Table.RenameColumns(ComMes, {
        {"Mês Lançamento", "SG_MES_COMPLETO"},
        {"Unidade Orçamentária", "CO_UO"}, {"Unidade Orçamentária (nome)", "NO_UO"},
        {"UG Executora", "CO_UG"}, {"UG Executora (nome)", "NO_UG"},
        {"Resultado Primário Lei", "ID_RESULTADO_LEI"}, {"Resultado Primário Lei (nome)", "NO_RESULTADO_LEI"},
        {"Autor Emendas Orçamento", "NO_AUTOR_EMENDA"},
        {"Ação Governo", "ID_ACAO_PT"}, {"Ação Governo (nome)", "NO_ACAO_PT"},
        {"Plano Orçamentário", "ID_PO"}, {"Plano Orçamentário (nome)", "NO_PO"},
        {"Grupo Despesa", "ID_GRUPO_DESPESA_NADE"}, {"Grupo Despesa (nome)", "NO_GRUPO_DESPESA_NADE"}},
        MissingField.Ignore),
    ComUO = Table.AddColumn(Renome, "ID_UO", each [CO_UO]),
    ComUO0 = Table.AddColumn(ComUO, "ID_UO0", each [CO_UO]),
    ComExec = Table.AddColumn(ComUO0, "ID_UG_EXEC", each [CO_UG]),
    ComNulos = List.Accumulate(
        {"ID_AUTOR_EMENDA", "NO_ACAO_PT", "NO_PO", "ID_FUNCAO_PT", "ID_SUBFUNCAO_PT", "ID_PROGRAMA_PT"}, ComExec,
        (t, c) => if List.Contains(Table.ColumnNames(t), c) then t
                  else Table.AddColumn(t, c, each null)),
    #"Cabeçalhos Promovidos" = Table.RemoveColumns(ComNulos, {"MET", "VALOR"}, MissingField.Ignore),
    // ========= DAQUI PARA BAIXO: PASSOS ORIGINAIS (inalterados) =========
    #"Tipo Alterado" = Table.TransformColumnTypes(#"Cabeçalhos Promovidos",{{"ID_UO", type text}, {"CO_UO", type text}, {"NO_UO", type text}, {"ID_UG_EXEC", type text}, {"CO_UG", type text}, {"NO_UG", type text}, {"ID_RESULTADO_LEI", type text}, {"NO_RESULTADO_LEI", type text}, {"ID_AUTOR_EMENDA", type text}, {"NO_AUTOR_EMENDA", type text}, {"ID_ACAO_PT", type text}, {"NO_ACAO_PT", type text}, {"ID_UO0", type text}, {"ID_FUNCAO_PT", type text}, {"ID_SUBFUNCAO_PT", type text}, {"ID_PROGRAMA_PT", type text}, {"ID_PO", type text}, {"NO_PO", type text}, {"ID_GRUPO_DESPESA_NADE", type text}, {"NO_GRUPO_DESPESA_NADE", type text}, {"ID_ITEM_INFORMACAO", type text}, {"NO_ITEM_INFORMACAO", type text}, {"CO_ITEM_INFORMACAO", type text}, {"SALDORCONTACONTBIL", Currency.Type}}),
    #"Texto Extraído Antes do Delimitador" = Table.TransformColumns(#"Tipo Alterado", {{"NO_AUTOR_EMENDA", each Text.BeforeDelimiter(_, " /"), type text}}),
    #"Coluna Mesclada Inserida" = Table.AddColumn(#"Texto Extraído Antes do Delimitador", "Grupo de Natureza de Despesa", each Text.Combine({[ID_GRUPO_DESPESA_NADE], [NO_GRUPO_DESPESA_NADE]}, " - "), type text),
    #"Valor Substituído" = Table.ReplaceValue(#"Coluna Mesclada Inserida","0","0000",Replacer.ReplaceText,{"ID_PO"}),
    #"Coluna Mesclada Inserida1" = Table.AddColumn(#"Valor Substituído", "Tipo de Emenda", each Text.Combine({[ID_RESULTADO_LEI], [NO_RESULTADO_LEI]}, " - "), type text),
    #"Coluna dividida" = Table.TransformColumns(#"Coluna Mesclada Inserida1", {{"SALDORCONTACONTBIL", each _ / 1000000, Currency.Type}})
in
    #"Coluna dividida"