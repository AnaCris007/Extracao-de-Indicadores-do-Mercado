"""Interface visual pra gerar a planilha de premissas macro escolhendo
indicadores e período numa tela, em vez de editar o script.

Roda em cima das mesmas funções de coleta de `premissas_macro.py` - não
duplica lógica, só monta a tela em volta de `montar_tabela()`.

pip install streamlit pandas requests yfinance openpyxl

Uso:
    streamlit run scripts/app.py
"""

from datetime import date

import streamlit as st

from premissas_macro import (
    ANO_FIM_PADRAO,
    ANO_INICIO_PADRAO,
    OUTPUT,
    TODOS_INDICADORES,
    formatar_planilha,
    montar_tabela,
)

st.set_page_config(page_title="Premissas Macro", page_icon="📊", layout="wide")
st.title("Premissas macro para valuation")
st.caption("IPCA, Selic, PIB, Ouro e Prata — fontes oficiais: IBGE, Banco Central e Yahoo Finance")

with st.form("filtros"):
    indicadores = st.multiselect("Indicadores", TODOS_INDICADORES, default=TODOS_INDICADORES)

    col1, col2 = st.columns(2)
    ano_inicio = col1.number_input("Ano inicial (histórico)", min_value=2010, max_value=date.today().year, value=ANO_INICIO_PADRAO)
    ano_fim = col2.number_input("Ano final (projeção)", min_value=date.today().year, max_value=date.today().year + 10, value=ANO_FIM_PADRAO)

    gerar = st.form_submit_button("Gerar planilha", type="primary")

if gerar:
    if not indicadores:
        st.warning("Selecione ao menos um indicador.")
    elif ano_inicio > ano_fim:
        st.warning("O ano inicial não pode ser depois do ano final.")
    else:
        with st.spinner("Buscando dados nas APIs oficiais (IBGE, BCB, Yahoo Finance)..."):
            df = montar_tabela(indicadores, int(ano_inicio), int(ano_fim))
            df.to_excel(OUTPUT, index=False, sheet_name="Premissas")
            formatar_planilha(OUTPUT, len(df))

        if df.empty:
            st.error("Nenhum dado retornado para essa combinação de indicadores/período.")
        else:
            st.success(f"{len(df)} linhas geradas para {len(indicadores)} indicador(es).")
            with open(OUTPUT, "rb") as f:
                st.download_button(
                    "⬇️ Baixar Excel",
                    f,
                    file_name="premissas_valuation.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.dataframe(df, use_container_width=True, height=420)
