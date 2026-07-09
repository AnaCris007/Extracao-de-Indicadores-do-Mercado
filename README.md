# Premissas macro para valuation

Coleta IPCA, Selic, PIB (histórico realizado + projeção Focus/BCB), Ouro
e Prata de fontes oficiais/gratuitas e gera uma planilha Excel formatada.

```bash
pip install -r requirements.txt
```

## Linha de comando

Gera a planilha com todos os indicadores e o período padrão:

```bash
python scripts/premissas_macro.py
```

## Interface visual

Tela pra escolher quais indicadores e qual período (ano inicial/final)
incluir, sem editar código:

```bash
streamlit run scripts/app.py
```

Abre em `http://localhost:8501`. Cada pessoa que rodar vê a mesma tela,
escolhe sua própria combinação e baixa o `.xlsx` gerado.

## Saída

`data/premissas_valuation.xlsx`, colunas `Data | Indicador | Tipo | Valor
| Unidade | AnoReferencia`, cabeçalho formatado e cor por `Tipo`
(Histórico/Projeção).

Detalhes de cada fonte estão comentados no topo de
[`scripts/premissas_macro.py`](scripts/premissas_macro.py).
