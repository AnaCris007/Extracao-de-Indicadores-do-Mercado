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
