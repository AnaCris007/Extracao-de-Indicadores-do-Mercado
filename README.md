# Premissas macro para valuation

Coleta IPCA, Selic, PIB (histórico realizado + projeção Focus/BCB), Ouro,
Prata e câmbio USD/BRL de fontes oficiais/gratuitas e gera uma planilha
Excel formatada, pronta pra alimentar premissas de um modelo de valuation.

**App publicado:** https://extracao-de-indicadores-do-mercado.streamlit.app/

Abre no navegador, sem instalar nada — escolhe os indicadores e o período,
clica em gerar e baixa o `.xlsx`.

## Como funciona

O projeto é dividido em duas camadas:

1. **Coleta e montagem** ([`scripts/premissas_macro.py`](scripts/premissas_macro.py))
   — funções que buscam cada indicador na fonte oficial e devolvem tabelas
   no mesmo formato longo (`Data, Indicador, Tipo, Valor, Unidade,
   AnoReferencia`). `montar_tabela(indicadores, ano_inicio, ano_fim)` junta
   tudo e filtra pelo período pedido; `pivotar_tabela(df)` transforma esse
   resultado no formato largo (indicador x ano) que vai pra planilha final.
2. **Interface** — duas formas de rodar essa coleta:
   - [`scripts/app.py`](scripts/app.py): tela web (Streamlit) pra escolher
     indicadores/período com cliques, sem editar código.
   - CLI (`python scripts/premissas_macro.py`): gera a planilha com todos
     os indicadores e o período padrão, direto no terminal.

Para cada indicador macro (IPCA, Selic, PIB), a lógica é a mesma: puxa o
valor **realizado** (fechado, oficial) ano a ano até o último ano
disponível, e completa o resto com a **projeção** — a mediana mais
recente da pesquisa Focus do Banco Central — até o ano final escolhido.
Isso evita lacunas e evita misturar dado real com estimativa sem deixar
claro qual é qual (coluna `Tipo`).

Ouro e Prata seguem cotação diária (não têm projeção), em **USD por onça
troy** — como são negociados no mercado internacional, sem conversão pra
real. USD/BRL é diferente dos demais: só traz a **cotação mais atual**
(PTAX do Banco Central), 1 valor só, no ano corrente — não série
histórica.

## Como cada indicador é calculado

Detalhe função a função (todas em
[`scripts/premissas_macro.py`](scripts/premissas_macro.py)):

### IPCA

- **Histórico** — [`get_ipca_historico()`](scripts/premissas_macro.py): IBGE/SIDRA,
  tabela 1737, variável `69` ("variação acumulada no ano"). A API devolve
  1 linha por mês; o código filtra só os meses de **dezembro**
  (`D3C.endswith("12")`), porque o acumulado de dezembro *é* o IPCA
  fechado do ano inteiro. Anos sem dezembro publicado ainda (`V == "..."`)
  são descartados.
- **Trimestral** — [`get_ipca_trimestral(ano)`](scripts/premissas_macro.py): mesma
  tabela, variável `2263` ("variação acumulada em 3 meses" — janela
  móvel). Filtrando só os meses `03/06/09/12`, essa janela móvel coincide
  exatamente com o trimestre civil (jan-fev-mar, abr-mai-jun, ...), sem
  precisar somar/compor os 3 meses manualmente. Só usada pra quebrar a
  coluna do ano corrente (ver seção de trimestres abaixo).
- **Projeção** — mediana Focus/BCB, indicador `"IPCA"` (ver "Projeção
  Focus" mais abaixo), pros anos sem dezembro fechado.

### Selic

- **Histórico** — [`get_selic_historico(ano_inicio)`](scripts/premissas_macro.py):
  Banco Central, SGS série `432` (meta Selic diária). Essa série repete o
  valor vigente em todo dia útil (só muda em reunião do Copom); o código
  agrupa por ano e pega o **último valor do ano** (`.groupby(ano).last()`)
  — a Selic com que aquele ano realmente fechou. Só inclui anos
  **anteriores ao corrente** (`Data.dt.year < HOJE.year`): o ano corrente
  em diante vem sempre da projeção Focus, mesmo que já tenha dado
  parcial do ano.
- **Projeção** — mediana Focus/BCB, indicador `"Selic"`. Vale notar: o
  Focus tem um indicador dedicado por *reunião* do Copom
  (`ExpectativasMercadoSelic`) e não publica pesquisa **mensal** de
  Selic (só existe pesquisa mensal pra IPCA e Câmbio) — então "Selic
  média do ano", quando alguém pede isso, na prática é a mesma mediana
  **anual** que o código já busca aqui (confirmado comparando com a
  aba Premissas do modelo da Vivara: os números batem exatamente).

### PIB

- **Histórico** — [`get_pib_historico()`](scripts/premissas_macro.py): IBGE/SIDRA,
  tabela 6784 (Contas Nacionais Trimestrais/Anuais), variável `9810`
  (variação real anual). O IBGE fecha essa série com defasagem de ~1-2
  anos em relação a IPCA/Selic — por isso é normal faltar o(s) ano(s)
  mais recente(s) aqui, viram projeção Focus automaticamente.
- **Projeção** — mediana Focus/BCB, indicador `"PIB Total"`.

### Projeção Focus (IPCA/Selic/PIB)

[`get_focus(indicador, nome, unidade, anos)`](scripts/premissas_macro.py) busca, pra
cada ano em `anos`, a **pesquisa mais recente** do Boletim Focus/BCB
(Olinda OData, `ExpectativasMercadoAnuais`), filtrando
`Indicador eq '<nome>' and DataReferencia eq '<ano>' and baseCalculo eq 0`
e ordenando por `Data desc` (`$top=1`) — ou seja, sempre a mediana mais
nova disponível pra aquele ano-alvo, não o histórico de como a previsão
mudou ao longo do tempo. Quais anos entram nessa lista quem decide é
[`anos_sem_historico(historico, ano_inicio, ano_fim)`](scripts/premissas_macro.py):
tudo que sobrou depois do último ano fechado do histórico até o fim do
período pedido. As chamadas — 1 requisição HTTP por ano — rodam em
paralelo (`ThreadPoolExecutor`), timeout curto (15s) e só 2 tentativas:
se a API do BCB estiver fora do ar (comum atrás de firewall
corporativo/institucional), é melhor pular aquele ano com um aviso do
que travar a geração inteira por minutos.

### Ouro e Prata

[`get_commodity(ticker, nome, ano_inicio)`](scripts/premissas_macro.py) usa a
biblioteca `yfinance` (`yf.Ticker("GC=F").history(...)` /
`yf.Ticker("SI=F").history(...)`) pra pegar o fechamento diário desde
`ano_inicio`. GC=F/SI=F são os futuros de ouro/prata da COMEX, cotados
nativamente em **USD por onça troy** — o código não converte pra real,
mantém a cotação como negociada no mercado internacional. Se o Yahoo
Finance falhar de forma transitória (acontece, sem lançar exceção — só
devolve uma tabela vazia), tenta de novo uma vez antes de desistir e
pular o indicador.

### USD/BRL

[`get_usd_brl()`](scripts/premissas_macro.py) é o único indicador que **não**
busca série histórica: chama `/dados/ultimos/1` da série SGS `1`
(câmbio PTAX venda) do Banco Central, que devolve só a cotação mais
recente disponível — 1 linha só, carimbada no ano corrente. Decisão
deliberada: como o objetivo aqui é o índice do dia (não uma série pra
gráfico), baixar anos de série diária só pra usar o último valor seria
desperdício de chamada.

### Como os anos/trimestres viram planilha

- [`montar_tabela(indicadores, ano_inicio, ano_fim)`](scripts/premissas_macro.py):
  chama a função de coleta certa pra cada indicador pedido, concatena
  tudo (histórico + projeção) numa tabela só no formato longo (`Data,
  Indicador, Tipo, Valor, Unidade, AnoReferencia`) e filtra pro período
  pedido.
- [`pivotar_tabela(df)`](scripts/premissas_macro.py): agrupa por
  `Indicador + AnoReferencia` e pega o **último valor por data**
  (`.last()`) de cada grupo — pra indicadores anuais (IPCA/Selic/PIB) só
  existe 1 valor por grupo mesmo; pra indicadores diários (Ouro, Prata),
  isso automaticamente vira "o valor de fechamento do ano" (o último dia
  cotado). Depois `.unstack("AnoReferencia")` transforma isso na matriz
  indicador × ano que vai pra planilha.
- [`adicionar_trimestres(valores, tipos, df, ano)`](scripts/premissas_macro.py):
  só age sobre o ano corrente (`ANO_TRIMESTRAL`), quebrando a coluna
  dele em `1T | 2T | 3T | 4T`. Pra Ouro/Prata, agrupa a série diária já
  coletada por trimestre (`Data.dt.quarter`) e pega o último valor de
  cada um — sem nenhuma chamada extra. Pra USD/BRL, a única cotação
  coletada cai no trimestre da sua própria data. Pra IPCA, chama
  `get_ipca_trimestral` à parte. Selic e PIB ficam em branco nos
  trimestres — não há fonte trimestral pra eles hoje.
- [`formatar_planilha(...)`](scripts/premissas_macro.py): não calcula nada,
  só estiliza o `.xlsx` já salvo — cor por `Tipo` (verde=Histórico,
  laranja=Projeção), cabeçalho preto nas colunas trimestrais, formato
  numérico por linha (`%`, `USD` ou `R$`, decidido em
  `formato_numero(unidade)` pela `Unidade` de cada indicador).

## Tecnologias

| Camada | Ferramenta | Papel |
|---|---|---|
| Coleta de dados | [`requests`](https://requests.readthedocs.io/) | chamadas HTTP diretas às APIs do IBGE e do Banco Central |
| Cotações | [`yfinance`](https://pypi.org/project/yfinance/) | histórico de Ouro (`GC=F`) e Prata (`SI=F`), em USD, via Yahoo Finance |
| Dados tabulares | [`pandas`](https://pandas.pydata.org/) | limpeza, junção e filtro das séries |
| Planilha | [`openpyxl`](https://openpyxl.readthedocs.io/) | geração e formatação do `.xlsx` (cores, largura de coluna, centralização) |
| Interface web | [`streamlit`](https://streamlit.io/) | tela interativa e hospedagem gratuita (Streamlit Community Cloud) |

**Fontes de dados** (todas oficiais, públicas e gratuitas, sem chave de API):

| Indicador | Fonte | Detalhe |
|---|---|---|
| IPCA (histórico) | IBGE / SIDRA, tabela 1737 | variação acumulada no ano, 1 valor por ano fechado |
| Selic (histórico) | Banco Central, SGS série 432 | meta Copom realizada, valor de fechamento de cada ano |
| PIB (histórico) | IBGE / SIDRA, tabela 6784 | variação real anual (Contas Nacionais) — sai com defasagem de ~1-2 anos |
| IPCA / Selic / PIB (projeção) | Banco Central, Focus (Olinda OData) | mediana da pesquisa mais recente por ano-alvo |
| Ouro / Prata | Yahoo Finance (`yfinance`) | fechamento diário, em USD/onça troy, sem conversão |
| USD/BRL | Banco Central, SGS série 1 (`/dados/ultimos/1`) | câmbio PTAX (venda), só a cotação mais atual |

## Rodando localmente

```bash
pip install -r requirements.txt
```

### Interface visual

```bash
streamlit run scripts/app.py
```

Abre em `http://localhost:8501`. Escolhe indicadores e período (ano
inicial/final), clica em "Gerar planilha" e baixa o `.xlsx`.

### Linha de comando

Gera a planilha com todos os indicadores e o período padrão:

```bash
python scripts/premissas_macro.py
```

## Saída

`data/premissas_valuation.xlsx` em formato largo: **1 linha por
indicador, colunas `Unidade` e `Fonte`, e depois 1 coluna por ano** — ano
mais antigo à esquerda, mais recente à direita. Cada célula é colorida
por `Tipo` (verde = Histórico, laranja = Projeção) e formatada de acordo
com a unidade da linha (`%` pra IPCA/Selic/PIB, `USD` pra Ouro/Prata,
`R$` pro câmbio USD/BRL).

A coluna `Fonte` identifica de onde vem o valor daquela linha - nos
indicadores macro (IPCA, Selic, PIB), histórico e projeção vêm de fontes
diferentes (IBGE/BCB vs. Focus/BCB), por isso os dois aparecem juntos.

Indicadores anuais (IPCA, Selic, PIB) já têm um valor por ano; Ouro e
Prata (diários) usam o último valor observado em cada ano — mesmo
critério de "ano fechado" usado no histórico da Selic. USD/BRL só
preenche a coluna do ano corrente, com a cotação do dia.

### Colunas trimestrais do ano corrente

A coluna do ano corrente (`ANO_TRIMESTRAL` em `premissas_macro.py`, hoje
2026) vem quebrada em `1T-2026 | 2T-2026 | 3T-2026 | 4T-2026` antes da
coluna anual, com destaque visual (cabeçalho preto) - só pros indicadores
com dado sub-anual disponível:

- **IPCA**: acumulado de cada trimestre civil (IBGE/SIDRA), preenchido só
  pros trimestres já fechados.
- **Ouro / Prata**: último valor observado em cada trimestre, a partir da
  mesma série diária já coletada.
- **USD/BRL**: como só a cotação atual é buscada (não o histórico), ela
  cai no trimestre correspondente à data de hoje; os demais ficam em
  branco.
- **Selic / PIB**: sem quebra trimestral - as fontes atuais não têm
  fechamento nem projeção por trimestre.

Detalhes de cada fonte estão comentados no topo de
[`scripts/premissas_macro.py`](scripts/premissas_macro.py).
