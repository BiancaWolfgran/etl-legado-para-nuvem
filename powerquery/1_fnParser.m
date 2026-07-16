// ================================================================
// fnParser — cole numa Consulta Nula (Editor Avançado) e renomeie
// a consulta para exatamente: fnParser
// Lê um export bruto do sistema oficial de relatórios — EXCEL (.xlsx) ou CSV —
// detectando automaticamente formato, encoding (Win-1252 / UTF-8) e
// delimitador (vírgula / ponto-e-vírgula / tab).
// CSVs do sistema oficial começam com linhas de título de UMA célula; por isso
// a leitura usa largura fixa de 250 colunas — o Csv.Document inferiria
// a largura pela primeira linha e decapitaria o arquivo — e depois
// corta as colunas vazias sobrando à direita.
// Devolve a tabela "wide": dimensões viram "Nome" e "Nome (nome)";
// métricas pivotadas viram "MET|codigo|nome"; preâmbulo e linha
// Total descartados.
// ================================================================
(conteudo as binary, delimitador as text) as table =>
let
    Ancoras = {"Mês Lançamento", "Ação Governo", "Unidade Orçamentária", "UG Executora"},
    PrimeiraCelula = (t as table) as function =>
        let c0 = Table.ColumnNames(t){0}
        in (linha) => Text.Trim(Text.From(Record.Field(linha, c0) ?? "")),
    AchaCabecalho = (t as table) as table =>
        let f = PrimeiraCelula(t)
        in Table.Skip(t, each not List.Contains(Ancoras, f(_))),

    // 1. tenta como Excel
    TentaXlsx = try
        let
            Wb = Excel.Workbook(conteudo, null, true),
            Planilhas = Table.SelectRows(Wb, each [Kind] = "Sheet")
        in
            Planilhas{0}[Data],

    // 2. se não for Excel, testa combinações de encoding x delimitador,
    //    sempre com largura FIXA de 250 colunas
    Delims = List.Distinct({delimitador, ",", ";", "#(tab)"}),
    Encodings = {1252, 65001},
    Combos = List.Combine(List.Transform(Encodings,
        (e) => List.Transform(Delims, (d) => [enc = e, delim = d]))),
    LerCsv = (c as record) as table =>
        Csv.Document(conteudo,
            [Delimiter = c[delim], Columns = 250,
             Encoding = c[enc], QuoteStyle = QuoteStyle.Csv]),
    ComboValido = List.First(
        List.Select(Combos, (c) =>
            (try Table.RowCount(AchaCabecalho(LerCsv(c))) > 2 otherwise false)),
        null),

    Bruto = if not TentaXlsx[HasError] then TentaXlsx[Value]
        else if ComboValido <> null then LerCsv(ComboValido)
        else error Error.Record("fnParser",
            "Cabeçalho do sistema oficial não encontrado no arquivo. " &
            "Confira se é o export correto: deve ter a coluna 'Mês Lançamento', " &
            "'Ação Governo' ou 'Unidade Orçamentária' na primeira posição.", null),

    SemPreambulo = AchaCabecalho(Bruto),
    H1cheio = Record.ToList(SemPreambulo{0}),
    H2cheio = Record.ToList(SemPreambulo{1}),

    // corta as colunas vazias à direita: a largura real vai até o último
    // cabeçalho não-vazio da primeira linha (a última coluna é sempre métrica)
    NaoVazio = (v) => Text.Trim(Text.From(v ?? "")) <> "",
    UltimoIdx = List.Last(List.Select(List.Positions(H1cheio), (i) => NaoVazio(H1cheio{i}))),
    n = UltimoIdx + 1,
    H1 = List.FirstN(H1cheio, n),
    H2 = List.FirstN(H2cheio, n),
    Dados = Table.SelectColumns(Table.Skip(SemPreambulo, 2),
        List.FirstN(Table.ColumnNames(SemPreambulo), n)),

    EhNumero = (t as text) as logical =>
        t <> "" and Text.Select(t, {"0".."9"}) = t,
    Nomes = List.Accumulate({0 .. n - 1}, {},
        (acc, i) =>
            let
                h = Text.Trim(Text.From(H1{i} ?? "")),
                m = Text.Trim(Text.From(H2{i} ?? "")),
                nome =
                    if h = "" then Text.Trim(Text.From(H1{i - 1} ?? "")) & " (nome)"
                    else if EhNumero(h) then "MET|" & h & "|" & m
                    else h
            in
                acc & {nome}),
    Renomeada = Table.RenameColumns(Dados,
        List.Zip({Table.ColumnNames(Dados), Nomes})),
    SemTotal = Table.SelectRows(Renomeada,
        each Text.Trim(Text.From(Record.Field(_, Nomes{0}) ?? "")) <> "Total")
in
    SemTotal
