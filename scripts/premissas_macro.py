"""Planilha de premissas macro para valuation: IPCA, Selic, PIB, Ouro e Prata.

Fontes oficiais/gratuitas, sem chave de API:
    - IPCA, histórico              -> IBGE / SIDRA (tabela 1737, variação
                                      acumulada no ano), 1 linha por ano fechado
    - Selic, histórico             -> Banco Central (SGS série 432, meta Selic
                                      realizada), 1 linha por ano fechado
    - PIB, histórico                -> IBGE / SIDRA (tabela 6784, Contas
                                      Nacionais - variação real anual), 1
                                      linha por ano já publicado pelo IBGE
                                      (esse dado sai com defasagem de ~1-2
                                      anos - é normal faltar o(s) ano(s) mais
                                      recente(s), viram projeção até lá)
    - IPCA/Selic/PIB, projeção     -> Focus / Banco Central (Olinda OData),
                                      mediana da pesquisa mais recente, 1
                                      linha por ano ainda sem dado oficial
                                      fechado (curva de projeção, não o
                                      histórico de como a previsão mudou)
    - Ouro e Prata, histórico      -> Yahoo Finance (tickers GC=F e SI=F),
                                      convertidos de USD para BRL pelo
                                      câmbio PTAX oficial do Banco Central

As funções de coleta (`get_*`) e `montar_tabela()` também são reaproveitadas
pelo app visual em `scripts/app.py` (`streamlit run scripts/app.py`), que
deixa escolher indicadores/período numa tela em vez de editar o código.

pip install pandas requests yfinance openpyxl

Uso:
    python scripts/premissas_macro.py
"""

from datetime import date, timedelta
from urllib.parse import quote

import pandas as pd
import requests
import yfinance as yf
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HOJE = date.today()
ANO_INICIO_PADRAO = HOJE.year - 7  # janela padrão de histórico
ANO_FIM_PADRAO = HOJE.year + 4  # até onde a projeção Focus vai por padrão
OUTPUT = "data/premissas_valuation.xlsx"

# Indicadores "macro" (1 linha/ano, Histórico até fechar + Projeção Focus
# depois disso) e o nome que a API do BCB Focus usa para cada um.
INDICADORES_MACRO = {
    "IPCA": "% a.a. (variação acumulada no ano)",
    "Selic": "% a.a. (taxa básica de juros)",
    "PIB": "% a.a. (variação do PIB no ano)",
}
INDICADOR_FOCUS = {"IPCA": "IPCA", "Selic": "Selic", "PIB": "PIB Total"}

# Commodities (histórico diário, sem projeção) e seu ticker na Yahoo Finance.
INDICADORES_COMMODITY = {"Ouro": "GC=F", "Prata": "SI=F"}

TODOS_INDICADORES = list(INDICADORES_MACRO) + list(INDICADORES_COMMODITY)


def get_ipca_historico() -> pd.DataFrame:
    """IPCA realizado, variação acumulada no ano (%), 1 linha por ano - IBGE/SIDRA.

    Variável 69 da tabela 1737 é o acumulado desde janeiro; pegando só o
    valor de dezembro de cada ano, obtemos o IPCA fechado do ano - no
    mesmo formato (% a.a.) da projeção do Focus, pra ficar comparável.
    """
    url = "https://apisidra.ibge.gov.br/values/t/1737/n1/all/v/69/p/all"
    rows = requests.get(url, timeout=30).json()[1:]  # 1ª linha é cabeçalho
    df = pd.DataFrame(rows)
    df = df[df["D3C"].str.endswith("12") & (df["V"] != "...")]  # só dezembro = ano fechado
    return pd.DataFrame(
        {
            "Data": pd.to_datetime(df["D3C"], format="%Y%m"),
            "Indicador": "IPCA",
            "Tipo": "Histórico",
            "Valor": df["V"].astype(float),
            "Unidade": INDICADORES_MACRO["IPCA"],
            "AnoReferencia": df["D3C"].str[:4].astype(int),
        }
    ).reset_index(drop=True)


def get_pib_historico() -> pd.DataFrame:
    """PIB realizado, variação real anual (%), 1 linha por ano - IBGE/SIDRA
    (Contas Nacionais Anuais, tabela 6784).

    O IBGE fecha as Contas Nacionais Anuais com defasagem de ~1-2 anos, ao
    contrário do IPCA/Selic - então o(s) ano(s) mais recente(s) simplesmente
    não aparecem aqui ainda; `montar_tabela()` completa a lacuna com a
    projeção Focus.
    """
    url = "https://apisidra.ibge.gov.br/values/t/6784/n1/all/v/9810/p/all"
    rows = requests.get(url, timeout=30).json()[1:]  # 1ª linha é cabeçalho
    df = pd.DataFrame(rows)
    df = df[df["V"] != "..."]
    return pd.DataFrame(
        {
            "Data": pd.to_datetime(df["D3C"] + "-12-31"),
            "Indicador": "PIB",
            "Tipo": "Histórico",
            "Valor": df["V"].astype(float),
            "Unidade": INDICADORES_MACRO["PIB"],
            "AnoReferencia": df["D3C"].astype(int),
        }
    ).reset_index(drop=True)


def get_sgs_serie(codigo: int, ano_inicio: int) -> pd.DataFrame:
    """Busca uma série diária do SGS/BCB (colunas Data/Valor), quebrando a
    consulta em janelas de até 10 anos.

    A própria API do Banco Central recusa (HTTP 406) qualquer consulta de
    série diária com janela maior que 10 anos - então pedir, por exemplo,
    2015 a 2030 numa chamada só falha. Aqui isso é escondido do resto do
    código: quem chama só pensa em "ano inicial", sem se preocupar com o
    limite.
    """
    url = f"https://api.bcb.gov.br/dados/serie/bcdata.sgs.{codigo}/dados"
    inicio = date(ano_inicio, 1, 1)
    partes = []
    while inicio <= HOJE:
        fim_janela = min(inicio + timedelta(days=3650), HOJE)  # ~9 anos e 11 meses, sempre < 10 anos
        params = {"formato": "json", "dataInicial": inicio.strftime("%d/%m/%Y"), "dataFinal": fim_janela.strftime("%d/%m/%Y")}

        payload = None
        for tentativa in range(2):  # a API do BCB às vezes falha/demora de forma transitória
            try:
                payload = requests.get(url, params=params, timeout=30).json()
                break
            except (requests.RequestException, ValueError):
                if tentativa == 1:
                    print(f"aviso: falha ao buscar série {codigo} de {inicio} a {fim_janela}, pulando janela")

        if isinstance(payload, list) and payload:
            partes.append(pd.DataFrame(payload))
        inicio = fim_janela + timedelta(days=1)

    if not partes:
        return pd.DataFrame(columns=["Data", "Valor"])
    df = pd.concat(partes, ignore_index=True)
    df["Data"] = pd.to_datetime(df["data"], format="%d/%m/%Y")
    df["Valor"] = df["valor"].astype(float)
    return df[["Data", "Valor"]].sort_values("Data").reset_index(drop=True)


def get_selic_historico(ano_inicio: int) -> pd.DataFrame:
    """Selic (meta Copom) realizada, 1 linha por ano já fechado - BCB, SGS série 432.

    A série 432 repete o valor vigente em todo dia útil (a meta só muda em
    reunião do Copom); aqui pegamos o último valor de cada ano - a Selic
    com que aquele ano realmente fechou. Só anos anteriores ao corrente
    entram aqui: o ano corrente em diante já vem do Focus (get_focus), que
    é projeção por definição, ainda não fechou.
    """
    df = get_sgs_serie(432, ano_inicio)
    df = df[df["Data"].dt.year < HOJE.year]

    ultimo_por_ano = df.sort_values("Data").groupby(df["Data"].dt.year).last()
    return pd.DataFrame(
        {
            "Data": ultimo_por_ano["Data"],
            "Indicador": "Selic",
            "Tipo": "Histórico",
            "Valor": ultimo_por_ano["Valor"],
            "Unidade": INDICADORES_MACRO["Selic"],
            "AnoReferencia": ultimo_por_ano.index,
        }
    ).reset_index(drop=True)


def anos_sem_historico(historico: pd.DataFrame, ano_inicio: int, ano_fim: int) -> list[int]:
    """Anos entre o último dado oficial fechado e o fim da janela pedida.

    O IPCA/Selic fecham o ano anterior; o PIB do IBGE atrasa mais (só sai
    ~1-2 anos depois) - por isso cada indicador pode ter uma lacuna de
    tamanho diferente, calculada aqui em vez de fixa. Se não há histórico
    nenhum no período (ex.: usuário pediu só anos futuros), cobre tudo com
    projeção.
    """
    if historico.empty:
        return list(range(ano_inicio, ano_fim + 1))
    ultimo_ano_fechado = int(historico["AnoReferencia"].max())
    return list(range(max(ultimo_ano_fechado + 1, ano_inicio), ano_fim + 1))


def get_focus(indicador: str, nome: str, unidade: str, anos: list[int]) -> pd.DataFrame:
    """Curva de projeção Focus/BCB: só a pesquisa mais recente por ano-alvo.

    Uma linha por ano em `anos`, cada uma com a mediana da última pesquisa
    disponível pra aquele ano - não o histórico de como a previsão mudou ao
    longo do tempo (isso existe, mas não serve pra alimentar um modelo de
    valuation: o que importa é o consenso de mercado *atual* para cada ano
    futuro).
    """
    linhas = []
    for ano in anos:
        filtro = f"Indicador eq '{indicador}' and DataReferencia eq '{ano}' and baseCalculo eq 0"
        params = {"$filter": filtro, "$orderby": "Data desc", "$top": "1", "$format": "json"}
        # A API do BCB só aceita espaço como %20 (não '+'), por isso o encode manual.
        query = "&".join(f"{k}={quote(v)}" for k, v in params.items())
        url = f"https://olinda.bcb.gov.br/olinda/servico/Expectativas/versao/v1/odata/ExpectativasMercadoAnuais?{query}"
        payload = requests.get(url, timeout=30).json()["value"]
        if not payload:  # nenhuma pesquisa registrada pra esse ano ainda
            continue
        registro = payload[0]
        linhas.append(
            {
                "Data": pd.to_datetime(registro["Data"]),
                "Indicador": nome,
                "Tipo": "Projeção",
                "Valor": float(registro["Mediana"]),
                "Unidade": unidade,
                "AnoReferencia": ano,
            }
        )
    return pd.DataFrame(linhas)


def get_usd_brl(ano_inicio: int) -> pd.DataFrame:
    """Câmbio USD/BRL (PTAX venda) - Banco Central, série SGS 1."""
    df = get_sgs_serie(1, ano_inicio)
    return df.rename(columns={"Valor": "USDBRL"})


def get_commodity(ticker: str, nome: str, usd_brl: pd.DataFrame, ano_inicio: int) -> pd.DataFrame:
    """Cotação histórica de fechamento - Yahoo Finance, convertida para BRL.

    Diferente de get_ipca/get_focus, aqui não há uma URL explícita: a
    biblioteca `yfinance` monta e chama a API do Yahoo Finance
    internamente (endpoint privado, não documentado publicamente -
    https://query1.finance.yahoo.com/v8/finance/chart/<ticker>), então o
    HTTP fica escondido dentro de `yf.Ticker(ticker).history(...)`.

    GC=F e SI=F (futuros de ouro/prata da COMEX) são cotados nativamente
    em dólar por onça troy (1 onça troy = 31,1035 g). Aqui multiplicamos
    pelo câmbio PTAX do próprio dia (ou do último dia útil disponível, via
    merge_asof) pra converter para BRL/onça troy.
    """
    hist = yf.Ticker(ticker).history(start=date(ano_inicio, 1, 1), interval="1d")
    dates = hist.index.tz_localize(None).astype("datetime64[ns]")
    df = pd.DataFrame({"Data": dates, "ValorUSD": hist["Close"]}).sort_values("Data")
    df = pd.merge_asof(df, usd_brl.astype({"Data": "datetime64[ns]"}), on="Data", direction="backward")
    return pd.DataFrame(
        {
            "Data": df["Data"],
            "Indicador": nome,
            "Tipo": "Histórico",
            "Valor": (df["ValorUSD"] * df["USDBRL"]).round(2),
            "Unidade": "BRL/onça troy",
            "AnoReferencia": df["Data"].dt.year,
        }
    )


def montar_tabela(indicadores: list[str], ano_inicio: int, ano_fim: int) -> pd.DataFrame:
    """Coleta só os indicadores/anos pedidos e devolve a tabela consolidada.

    `indicadores` é qualquer subconjunto de `TODOS_INDICADORES`. Usada tanto
    pelo `main()` (CLI, roda com tudo) quanto pelo `app.py` (Streamlit, roda
    só com o que o usuário selecionou na tela).
    """
    tabelas: list[pd.DataFrame] = []
    usd_brl = None

    for nome in indicadores:
        if nome in INDICADORES_MACRO:
            historico = {
                "IPCA": get_ipca_historico,
                "Selic": lambda: get_selic_historico(ano_inicio),
                "PIB": get_pib_historico,
            }[nome]()
            anos = anos_sem_historico(historico, ano_inicio, ano_fim)
            projecao = get_focus(INDICADOR_FOCUS[nome], nome, INDICADORES_MACRO[nome], anos)
            tabelas += [historico, projecao]
        elif nome in INDICADORES_COMMODITY:
            if usd_brl is None:
                usd_brl = get_usd_brl(ano_inicio)
            tabelas.append(get_commodity(INDICADORES_COMMODITY[nome], nome, usd_brl, ano_inicio))

    if not tabelas:
        return pd.DataFrame(columns=["Data", "Indicador", "Tipo", "Valor", "Unidade", "AnoReferencia"])

    df = pd.concat(tabelas, ignore_index=True)
    df = df[(df["AnoReferencia"] >= ano_inicio) & (df["AnoReferencia"] <= ano_fim)]
    return df.sort_values(["Indicador", "Data"]).reset_index(drop=True)


def formatar_planilha(caminho: str, n_linhas: int) -> None:
    """Aplica formatação visual ao .xlsx já salvo: cabeçalho, cores por Tipo,
    largura de coluna e formato numérico por unidade (%, R$).

    Reaberto depois do `df.to_excel` de propósito - separa "gerar o dado"
    de "deixar bonito", em vez de misturar estilo com a lógica de coleta.
    """
    from openpyxl import load_workbook

    wb = load_workbook(caminho)
    ws = wb["Premissas"]

    col = {cell.value: cell.column for cell in ws[1]}  # nome da coluna -> índice
    cor_header = PatternFill("solid", fgColor="1F4E78")
    cor_historico = PatternFill("solid", fgColor="E2EFDA")  # verde claro
    cor_projecao = PatternFill("solid", fgColor="FCE4D6")  # laranja claro

    centralizado = Alignment(horizontal="center", vertical="center")

    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = cor_header
        cell.alignment = centralizado

    # Largura de cada coluna calculada a partir do texto que aparece na tela
    # (já formatado - "1.234,56 R$", "04/07/2026" -, não do valor cru).
    larguras: dict[int, int] = {}
    for row in range(2, n_linhas + 2):
        tipo = ws.cell(row, col["Tipo"]).value
        unidade = ws.cell(row, col["Unidade"]).value
        fill = cor_historico if tipo == "Histórico" else cor_projecao
        ws.cell(row, col["Tipo"]).fill = fill

        valor_cell = ws.cell(row, col["Valor"])
        eh_brl = "BRL" in str(unidade)
        valor_cell.number_format = '#,##0.00" R$"' if eh_brl else "0.00"
        texto_valor = (f"{valor_cell.value:,.2f} R$" if eh_brl else f"{valor_cell.value:.2f}") if valor_cell.value is not None else ""

        data_cell = ws.cell(row, col["Data"])
        data_cell.number_format = "dd/mm/yyyy"

        for nome_coluna, indice in col.items():
            cell = ws.cell(row, indice)
            cell.alignment = centralizado
            if nome_coluna == "Valor":
                texto = texto_valor
            elif nome_coluna == "Data":
                texto = data_cell.value.strftime("%d/%m/%Y") if data_cell.value else ""
            else:
                texto = str(cell.value) if cell.value is not None else ""
            larguras[indice] = max(larguras.get(indice, 0), len(texto))

    for nome_coluna, indice in col.items():
        largura_cabecalho = len(str(nome_coluna))
        largura = max(larguras.get(indice, 0), largura_cabecalho)
        ws.column_dimensions[get_column_letter(indice)].width = largura + 4  # levemente maior que o conteúdo

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    wb.save(caminho)


def main() -> None:
    df = montar_tabela(TODOS_INDICADORES, ANO_INICIO_PADRAO, ANO_FIM_PADRAO)
    df.to_excel(OUTPUT, index=False, sheet_name="Premissas")
    formatar_planilha(OUTPUT, len(df))
    print(f"Salvo em {OUTPUT} - {len(df)} linhas - indicadores: {sorted(df['Indicador'].unique())}")


if __name__ == "__main__":
    main()
