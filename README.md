# Premissas macro para valuation

Coleta IPCA, Selic, PIB (histórico realizado + projeção Focus/BCB), Ouro
e Prata de fontes oficiais/gratuitas e gera uma planilha Excel formatada,
pronta pra alimentar premissas de um modelo de valuation.

**App publicado:** https://extracao-de-indicadores-do-mercado.streamlit.app/

Abre no navegador, sem instalar nada — escolhe os indicadores e o período,
clica em gerar e baixa o `.xlsx`.

## Como funciona

O projeto é dividido em duas camadas:

1. **Coleta e montagem** ([`scripts/premissas_macro.py`](scripts/premissas_macro.py))
   — funções que buscam cada indicador na fonte oficial e devolvem tabelas
   no mesmo formato (`Data, Indicador, Tipo, Valor, Unidade, AnoReferencia`).
   `montar_tabela(indicadores, ano_inicio, ano_fim)` junta tudo, filtra pelo
   período pedido e formata a planilha final.
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

Ouro e Prata seguem cotação diária (não têm projeção), convertida de
dólar pra real pelo câmbio PTAX oficial do dia.

## Tecnologias

| Camada | Ferramenta | Papel |
|---|---|---|
| Coleta de dados | [`requests`](https://requests.readthedocs.io/) | chamadas HTTP diretas às APIs do IBGE e do Banco Central |
| Cotações | [`yfinance`](https://pypi.org/project/yfinance/) | histórico de Ouro (`GC=F`) e Prata (`SI=F`) via Yahoo Finance |
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
| Ouro / Prata | Yahoo Finance (`yfinance`) | fechamento diário, convertido pra BRL pelo câmbio PTAX (BCB, SGS série 1) |

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

`data/premissas_valuation.xlsx`, colunas `Data | Indicador | Tipo | Valor
| Unidade | AnoReferencia`, cabeçalho formatado, células centralizadas,
coluna `Tipo` colorida (verde = Histórico, laranja = Projeção) e largura
de coluna ajustada ao conteúdo.

Detalhes de cada fonte estão comentados no topo de
[`scripts/premissas_macro.py`](scripts/premissas_macro.py).
