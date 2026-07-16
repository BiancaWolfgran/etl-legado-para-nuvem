# Painel AREA/COORDENACAO 100% na nuvem — e-mail → SharePoint → Power BI
## (sem SQL, sem Python, sem provedor DaaS, sem máquina ligada)

## Arquitetura

```
sistema oficial de relatórios (agendamento por e-mail, já funciona)
        v
Power Automate  -> salva os 4 anexos no SharePoint com nome fixo (sobrescreve)
        v
SharePoint  .../CAMINHO/DA/PASTA/RELATORIOS/   (4 arquivos CSV, sempre os mesmos nomes)
        v
Power BI (Power Query lê e transforma direto do CSV bruto)
        v
Refresh agendado no Power BI Service (fonte na nuvem = nem gateway precisa)
```

---

## PASSO 1 — Criar a pasta no SharePoint

No site **SEU_SITE**, dentro de
`Documentos Compartilhados > SUA_PASTA > PLANILHAS_DOS_PAINEIS > sua estrutura de pastas`,
crie a pasta **RELATORIOS**.

Faça upload manual dos 4 CSVs de hoje com estes nomes EXATOS (o Power Query
vai procurar por eles):

| Relatório do sistema oficial           | Nome do arquivo na pasta  |
|--------------------------------|---------------------------|
| Execução completa              | `EXECUCAO_COMPLETA.xlsx`   |
| Execução por TED               | `EXECUCAO_TED.xlsx`        |
| Crédito indisponível           | `INDISPONIVEL.xlsx`        |
| Emendas                        | `EMENDAS.xlsx`             |

---

## PASSO 2 — Fluxo no Power Automate (salva os anexos do e-mail)

Em https://make.powerautomate.com > Criar > **Fluxo de nuvem automatizado**:

1. **Gatilho**: "Quando um novo email chegar (V3)" (Office 365 Outlook)
   - Em opções avançadas: filtre por remetente do sistema oficial de relatórios e/ou
     Assunto contém o nome do relatório; marque "Somente com anexos = Sim"
     e "Incluir anexos = Sim".
2. **Apply to each** sobre `Anexos` do gatilho. Dentro dele:
3. **Condição** (ou um "Alternar"/Switch) sobre `Nome do anexo`:
   - contém `EXECUCAO_COMPLETA`  -> ramo 1
   - contém `TED`                -> ramo 2
   - contém `INDISPONIVEL`       -> ramo 3
   - contém `EMENDAS`            -> ramo 4
4. Em cada ramo, duas ações do SharePoint:
   a. **Excluir arquivo** — Endereço do site: SEU_SITE;
      Arquivo: o caminho fixo (ex.: `/Documentos Compartilhados/SUA_PASTA/PLANILHAS_DOS_PAINEIS/CAMINHO/DA/PASTA/RELATORIOS/EXECUCAO_COMPLETA.xlsx`).
      IMPORTANTE: clique nos "..." da ação > **Configurar executar após** e, na
      ação seguinte (Criar arquivo), marque para executar mesmo se "Excluir
      arquivo" falhar — assim o fluxo não quebra na primeira execução, quando
      o arquivo ainda não existe.
   b. **Criar arquivo** — mesma pasta, Nome do arquivo fixo
      (ex.: `EXECUCAO_COMPLETA.xlsx`), Conteúdo do arquivo = `Conteúdo do anexo`.

Pronto: todo dia, quando os e-mails chegarem, os 4 arquivos são sobrescritos.
(Se os 4 relatórios chegarem em e-mails separados, o mesmo fluxo resolve —
cada e-mail dispara o fluxo e só o ramo do anexo correspondente executa.)

---

## PASSO 3 — Trocar as consultas no Power BI Desktop (uma única vez)

Abra o PBIX > **Transformar dados**. Na pasta `powerquery/` deste pacote há
5 arquivos de texto:

1. `1_fnParser.m` — Página Inicial > Nova Fonte > **Consulta Nula**.
   Abra o **Editor Avançado**, apague o que estiver lá, cole o conteúdo do
   arquivo e renomeie a consulta para **fnParser** (exatamente assim).
   É a função que resolve o cabeçalho duplo do sistema oficial, descarta o preâmbulo
   e a linha Total — usada pelas 4 consultas.
2. Para cada consulta existente — **Base_Orcamentaria**, **Base_Transferencia**,
   **Indisponível** e **emendas** — clique com o botão direito > Editor
   Avançado, apague TUDO e cole o conteúdo do arquivo `.m` correspondente.
3. Feche e Aplique.

O que os novos códigos fazem: leem o CSV bruto da pasta RELATORIOS, fazem o
unpivot dos itens de informação/contas contábeis, recriam as colunas com os
nomes que o modelo espera (ID_PTRES, NO_PO, SALDORITEMINFORMAO...) e então
executam OS MESMOS passos originais da consulta (joins com PO/AÇ/PR/UG,
chave_ptres, divisão por milhão nas emendas etc.). Nada muda no modelo, nas
medidas ou nos visuais.

Na primeira atualização o Power BI vai pedir credenciais para
`https://SEU_TENANT.sharepoint.com` — escolha **Conta organizacional** e entre
com seu usuário ORGAO.

### Observações importantes
- Os meses 013/014 (apuração/encerramento do sistema contábil federal) são preservados.
- A consulta Indisponível busca os NOMES de Ação e PO no arquivo de execução
  completa (o relatório de indisponível só traz códigos) — necessários para
  os merges com as dimensões AÇ e PO.
- Função, Subfunção, Localizador, Órgão do PI continuam NULL (os relatórios
  não os trazem). Se um dia você acrescentar esses atributos nos relatórios
  do sistema oficial, me chame que ajusto o M para capturá-los (a fnParser já lê
  qualquer par código/descrição automaticamente).
- Se o sistema oficial mudar o LAYOUT do relatório (ordem/nomes de colunas), as
  consultas quebram — basta reexportar no modelo combinado ou me chamar.

---

## PASSO 4 — Publicar e agendar

1. Página Inicial > **Publicar** no workspace de sempre.
2. No Power BI Service > Conjunto de dados > Configurações:
   - Em **Credenciais da fonte de dados**: SharePoint/Web -> OAuth2 /
     Conta organizacional.
   - Em **Atualização agendada**: ative e defina o horário (depois do horário
     em que os e-mails do sistema oficial costumam chegar; dá para colocar mais de um
     horário por dia).
3. Como TODAS as fontes agora são nuvem (SharePoint), o **gateway não é mais
   necessário** — se o dataset estiver amarrado a um gateway antigo do SQL,
   remova a associação nas configurações.

---

## Resultado final

- Zero provedor DaaS (pode encerrar o contrato dos 50 milheiros).
- Zero SQL Server, zero scripts, zero máquina ligada.
- Cadeia inteira: sistema oficial agenda o e-mail -> Power Automate salva ->
  Power BI atualiza sozinho.
- Manutenção: nenhuma, enquanto o layout dos relatórios do sistema oficial não mudar.
