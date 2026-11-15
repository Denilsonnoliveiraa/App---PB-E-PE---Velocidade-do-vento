
# ============================================================
# OBSERVATÓRIO DOS VENTOS — PB E PE | 1991–2025
# Denilson de Oliveira Silva
# Jader da Silva Jale
# Silvio Fernando Alves Xavier
#
# Layout amarelo e dourado
# Contraste corrigido em TODOS os botões
# ============================================================

from pathlib import Path
import re
import unicodedata
import warnings

import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from scipy import stats
from statsmodels.tsa.stattools import adfuller, kpss, acf
from statsmodels.tsa.seasonal import STL
from statsmodels.stats.diagnostic import acorr_ljungbox


# ============================================================
# 1. CONFIGURAÇÃO
# ============================================================

st.set_page_config(
    page_title="Observatório dos Ventos | PB e PE",
    page_icon="🌬️",
    layout="wide",
    initial_sidebar_state="expanded"
)

BASE = Path(__file__).resolve().parent

EXCEL_NOME = "Velocidade_Vento_Diaria_PB_PE_1991_2025.xlsx"

OPCOES_EXCEL = [
    BASE / EXCEL_NOME,
    BASE / "dados" / EXCEL_NOME,
    BASE.parent / EXCEL_NOME
]

EXCEL = next(
    (p for p in OPCOES_EXCEL if p.is_file()),
    OPCOES_EXCEL[0]
)

OPCOES_ATLAS = [
    BASE / "ATLAS_VENTO_PB_PE",
    BASE.parent / "ATLAS_VENTO_PB_PE",
    BASE
]

ATLAS = next(
    (
        p for p in OPCOES_ATLAS
        if (p / "02_MAPAS_INDIVIDUAIS").is_dir()
    ),
    OPCOES_ATLAS[0]
)

PASTAS = {
    "Mapas individuais": ATLAS / "02_MAPAS_INDIVIDUAIS",
    "Recortes temporais": ATLAS / "03_RECORTES_TEMPORAIS",
    "Painel geral": ATLAS / "01_PAINEL_GERAL"
}

ESTADOS = ["Paraíba", "Pernambuco"]
CLASSES = ["Fraco", "Moderado", "Forte"]
ESTACOES = ["Verão", "Outono", "Inverno", "Primavera"]

PERIODOS = [
    "1991-1995", "1996-2000", "2001-2005",
    "2006-2010", "2011-2015",
    "2016-2020", "2021-2025"
]

# ============================================================
# 2. CORES
# ============================================================

FUNDO = "#FFF9ED"
PAPEL = "#FFFFFF"
TEXTO = "#211A10"
ESCURO = "#302719"
MARROM = "#70451C"
DOURADO = "#D89323"
AMARELO = "#F8DC76"

CORES_ESTADOS = {
    "Paraíba": "#D89323",
    "Pernambuco": "#70451C"
}

CORES_MARKOV = {
    "Fraco": "#F2D374",
    "Moderado": "#DFA12C",
    "Forte": "#865018"
}

CORES_SERIES = [
    "#A96A17", "#D89323", "#70451C",
    "#C75B27", "#76539C", "#54728A",
    "#8B4513"
]


# ============================================================
# 3. CSS — CORREÇÃO GLOBAL DE CONTRASTE
# ============================================================

st.markdown("""
<style>

/* FUNDO GERAL */

:root {
    color-scheme: light;
}

.stApp,
[data-testid="stMain"],
[data-testid="stAppViewContainer"] {
    background: #FFF9ED !important;
    color: #211A10 !important;
}

[data-testid="stMainBlockContainer"] {
    max-width: 1500px;
    padding-top: 1.6rem;
    padding-bottom: 3rem;
}

[data-testid="stMain"] h1,
[data-testid="stMain"] h2,
[data-testid="stMain"] h3,
[data-testid="stMain"] h4 {
    color: #392D18 !important;
    font-family: "Segoe UI", Arial, sans-serif;
    font-weight: 750;
}

[data-testid="stMain"] p,
[data-testid="stMain"] label {
    color: #211A10;
}

/* BARRA LATERAL */

[data-testid="stSidebar"],
[data-testid="stSidebarContent"] {
    background: #302719 !important;
}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] span {
    color: #FFFFFF !important;
}

[data-testid="stSidebar"] hr {
    border-color: #8A754E !important;
}

/* ======================================================
   BOTÕES NORMAIS, DOWNLOADS E BOTÕES DE LINKS
   CORREÇÃO PRINCIPAL DESTA VERSÃO
====================================================== */

[data-testid="stBaseButton-primary"],
[data-testid="stBaseButton-secondary"],
[data-testid="stBaseButton-tertiary"],
[data-testid="stBaseButton-primaryFormSubmit"],
[data-testid="stBaseButton-secondaryFormSubmit"],
[data-testid="stLinkButton"] a,
[data-testid="stLinkButton"] a:link,
[data-testid="stLinkButton"] a:visited,
.stButton > button,
.stDownloadButton > button,
.stFormSubmitButton > button,
.stLinkButton > a {
    background-color: #302719 !important;
    color: #FFFFFF !important;
    border: 1px solid #302719 !important;
    border-radius: 10px !important;
    font-weight: 700 !important;
    min-height: 43px;
    text-decoration: none !important;
    opacity: 1 !important;
}

/* Todos os textos dentro dos botões */

[data-testid="stButton"] button *,
[data-testid="stDownloadButton"] button *,
[data-testid="stLinkButton"] a *,
[data-testid="stFormSubmitButton"] button *,
.stButton button p,
.stDownloadButton button p,
.stLinkButton a p,
.stFormSubmitButton button p,
.stButton button span,
.stDownloadButton button span,
.stLinkButton a span {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
    opacity: 1 !important;
}

/* Hover */

[data-testid="stButton"] button:hover,
[data-testid="stDownloadButton"] button:hover,
[data-testid="stLinkButton"] a:hover,
[data-testid="stFormSubmitButton"] button:hover,
.stLinkButton > a:hover {
    background: #865018 !important;
    border-color: #865018 !important;
    color: #FFFFFF !important;
}

/* Hover e foco de todos os textos internos */

[data-testid="stButton"] button:is(:hover,:focus,:active) *,
[data-testid="stDownloadButton"] button:is(:hover,:focus,:active) *,
[data-testid="stLinkButton"] a:is(:hover,:focus,:active) *,
[data-testid="stFormSubmitButton"] button:is(:hover,:focus,:active) * {
    color: #FFFFFF !important;
    -webkit-text-fill-color: #FFFFFF !important;
}

/* Links desabilitados */

[data-testid="stLinkButton"] a[aria-disabled="true"],
[data-testid="stButton"] button:disabled {
    color: #FFFFFF !important;
    opacity: 0.65 !important;
}

/* ABAS */

button[data-baseweb="tab"] {
    background: #F6E7BB !important;
    padding: 11px 13px !important;
    border-radius: 8px 8px 0 0 !important;
    margin-right: 4px;
}

button[data-baseweb="tab"] * {
    color: #302719 !important;
    font-weight: 650 !important;
}

button[data-baseweb="tab"][aria-selected="true"] {
    background: #302719 !important;
}

button[data-baseweb="tab"][aria-selected="true"] * {
    color: #FFFFFF !important;
}

/* MÉTRICAS */

[data-testid="stMetric"] {
    background: #FFFFFF !important;
    border: 1px solid #E7D8AE;
    border-top: 4px solid #D89323;
    border-radius: 12px;
    padding: 15px;
}

[data-testid="stMetric"] * {
    color: #211A10 !important;
}

/* SELETORES */

[data-baseweb="select"] > div,
[data-baseweb="input"] > div {
    background: #FFFFFF !important;
    color: #211A10 !important;
    border-color: #E7D8AE !important;
}

[data-baseweb="popover"],
[data-baseweb="popover"] ul,
[data-baseweb="popover"] li {
    background: #FFFFFF !important;
    color: #211A10 !important;
}

[data-baseweb="popover"] * {
    color: #211A10 !important;
}

/* CARDS E EXPANSORES */

[data-testid="stVerticalBlockBorderWrapper"] {
    border-color: #E7D8AE !important;
    border-radius: 12px;
}

[data-testid="stExpander"] {
    background: #FFFFFF !important;
    border: 1px solid #E7D8AE;
    border-radius: 10px;
}

/* INDICADORES E GRÁFICOS */

[data-testid="stCaptionContainer"] {
    color: #665331;
}

[data-testid="stPlotlyChart"] {
    border-radius: 12px;
    overflow: hidden;
}

</style>
""", unsafe_allow_html=True)


# ============================================================
# 4. FUNÇÕES VISUAIS
# ============================================================

def grafico(fig, altura=430):
    fig.update_layout(
        template="plotly_white",
        height=altura,
        paper_bgcolor=PAPEL,
        plot_bgcolor=PAPEL,
        font=dict(color="#111111", size=12),
        title=dict(
            x=0.03,
            font=dict(color="#111111", size=17)
        ),
        legend=dict(
            bgcolor="#FFFFFF",
            font=dict(color="#111111", size=12),
            title=dict(
                font=dict(color="#111111")
            )
        ),
        hoverlabel=dict(
            bgcolor=ESCURO,
            font=dict(color="#FFFFFF")
        ),
        margin=dict(l=45, r=25, t=70, b=60)
    )

    fig.update_xaxes(
        showgrid=False,
        tickfont=dict(color="#111111"),
        title_font=dict(color="#111111")
    )

    fig.update_yaxes(
        gridcolor="#F0E6D4",
        tickfont=dict(color="#111111"),
        title_font=dict(color="#111111")
    )

    fig.update_coloraxes(
        colorbar=dict(
            tickfont=dict(color="#111111"),
            title=dict(font=dict(color="#111111"))
        )
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={
            "displaylogo": False,
            "toImageButtonOptions": {
                "format": "png",
                "scale": 3
            }
        }
    )


def numero(v, casas=2):
    if pd.isna(v):
        return "—"
    return f"{v:.{casas}f}".replace(".", ",")


def valor_p(v):
    if pd.isna(v):
        return "—"
    if v < 0.01:
        return "p < 0,01"
    return numero(v, 4)


def normalizar(valor):
    s = unicodedata.normalize(
        "NFKD",
        str(valor).strip().lower()
    )
    s = "".join(
        c for c in s
        if not unicodedata.combining(c)
    )
    return re.sub(
        r"[^a-z0-9]+", "_", s
    ).strip("_")


def para_numero(coluna):
    if pd.api.types.is_numeric_dtype(coluna):
        return pd.to_numeric(
            coluna, errors="coerce"
        )
    return pd.to_numeric(
        coluna.astype(str)
        .str.strip()
        .str.replace(",", ".", regex=False),
        errors="coerce"
    )


def baixar_csv(df, nome, chave):
    st.download_button(
        label="Baixar resultados CSV",
        data=df.to_csv(
            index=False
        ).encode("utf-8-sig"),
        file_name=nome,
        mime="text/csv",
        key=chave
    )


# ============================================================
# 5. CARREGAR DADOS DO EXCEL
# ============================================================

@st.cache_data(show_spinner="Carregando dados de vento...")
def ler_excel(caminho, modificacao):
    with pd.ExcelFile(caminho) as xls:
        abas = xls.sheet_names

        aba = next(
            (
                a for a in abas
                if normalizar(a) == "pb_pe_diario"
            ),
            None
        )

        if aba is None:
            raise ValueError(
                "A aba PB_PE_Diario não foi encontrada."
            )

        df = pd.read_excel(
            xls,
            sheet_name=aba
        )

    return df, abas


def preparar_dados(df):
    df = df.copy()

    df.columns = [
        normalizar(c)
        for c in df.columns
    ]

    if (
        "velocidade_vento_ms" not in df.columns
        and "velocidade_vento_m_s" in df.columns
    ):
        df = df.rename(columns={
            "velocidade_vento_m_s": "velocidade_vento_ms"
        })

    obrigatorias = [
        "data", "estado", "velocidade_vento_ms"
    ]

    faltantes = [
        c for c in obrigatorias
        if c not in df.columns
    ]

    if faltantes:
        raise ValueError(
            "Colunas ausentes: "
            + ", ".join(faltantes)
        )

    if pd.api.types.is_numeric_dtype(df["data"]):
        df["data"] = pd.to_datetime(
            df["data"],
            unit="D",
            origin="1899-12-30",
            errors="coerce"
        )
    else:
        df["data"] = pd.to_datetime(
            df["data"],
            errors="coerce"
        )

    df["data"] = df["data"].dt.normalize()

    df["vento"] = para_numero(
        df["velocidade_vento_ms"]
    )

    def reconhecer_estado(v):
        nome = normalizar(v)
        if nome == "paraiba":
            return "Paraíba"
        if nome == "pernambuco":
            return "Pernambuco"
        return None

    df["estado"] = df["estado"].apply(
        reconhecer_estado
    )

    df = df.dropna(
        subset=["data", "estado"]
    )

    df = (
        df.groupby(
            ["estado", "data"],
            as_index=False
        )["vento"].mean()
    )

    df = df.sort_values(
        ["estado", "data"]
    ).reset_index(drop=True)

    df["ano"] = df["data"].dt.year
    df["mes"] = df["data"].dt.month
    df["dia"] = df["data"].dt.day
    df["dia_do_ano"] = df["data"].dt.dayofyear

    inicio = (
        1991 + 5 * ((df["ano"] - 1991) // 5)
    )
    fim = np.minimum(inicio + 4, 2025)

    df["periodo"] = (
        inicio.astype(str)
        + "-"
        + fim.astype(str)
    )

    return df


dados = None
abas_excel = []
erro_excel = None

if EXCEL.is_file():
    try:
        bruto, abas_excel = ler_excel(
            str(EXCEL),
            EXCEL.stat().st_mtime_ns
        )
        dados = preparar_dados(bruto)

        if dados.empty:
            dados = None
            erro_excel = "A planilha está vazia."

    except Exception as erro:
        erro_excel = str(erro)
else:
    erro_excel = f"Planilha não encontrada: {EXCEL}"


# ============================================================
# 6. FUNÇÕES TEMPORAIS
# ============================================================

def serie_estado(df, estado):
    return (
        df.loc[
            df["estado"] == estado,
            ["data", "vento"]
        ]
        .set_index("data")["vento"]
        .sort_index()
        .asfreq("D")
    )


def serie_mensal(df, estado):
    return (
        serie_estado(df, estado)
        .resample("MS")
        .mean()
    )


def anomalias_mensais(df, estado):
    s = serie_mensal(df, estado)
    resultado = s.rename("vento").to_frame()

    resultado["mes"] = resultado.index.month
    resultado["referencia"] = (
        resultado.groupby("mes")["vento"]
        .transform("mean")
    )
    resultado["anomalia"] = (
        resultado["vento"]
        - resultado["referencia"]
    )
    resultado["ano"] = resultado.index.year

    return resultado.reset_index()


# ============================================================
# 7. FUNÇÕES DE MARKOV
# ============================================================

def quartis_vento(df):
    v = df["vento"].dropna()

    if len(v) < 20:
        raise ValueError("Poucos dados para quartis.")

    q1, q2, q3 = v.quantile(
        [0.25, 0.50, 0.75]
    ).to_numpy()

    if q1 >= q3:
        raise ValueError(
            "Q1 e Q3 não permitem três categorias."
        )

    return float(q1), float(q2), float(q3)


def classificar_vento(valores, q1, q3):
    classes = pd.Series(
        pd.NA,
        index=valores.index,
        dtype="object"
    )

    validos = valores.notna()

    classes.loc[
        validos & (valores <= q1)
    ] = "Fraco"

    classes.loc[
        validos
        & (valores > q1)
        & (valores <= q3)
    ] = "Moderado"

    classes.loc[
        validos & (valores > q3)
    ] = "Forte"

    return classes


def preparar_markov(df, estado, quartis):
    q1, q2, q3 = quartis

    serie = serie_estado(df, estado)

    tabela = pd.DataFrame({
        "data": serie.index,
        "vento": serie.values
    })

    tabela["classe"] = classificar_vento(
        tabela["vento"], q1, q3
    )

    tabela["ano"] = tabela["data"].dt.year
    tabela["mes"] = tabela["data"].dt.month

    return tabela


def matriz_markov(tabela):
    d = tabela.sort_values("data").copy()

    d["proxima_data"] = d["data"].shift(-1)
    d["proxima_classe"] = d["classe"].shift(-1)

    intervalo = (
        d["proxima_data"] - d["data"]
    ).dt.days

    trans = d.loc[
        (intervalo == 1)
        & d["classe"].notna()
        & d["proxima_classe"].notna()
    ]

    N = pd.crosstab(
        trans["classe"],
        trans["proxima_classe"]
    ).reindex(
        index=CLASSES,
        columns=CLASSES,
        fill_value=0
    )

    P = N.div(
        N.sum(axis=1).replace(0, np.nan),
        axis=0
    )

    return N, P, len(trans)


def estacionaria(P):
    if P.isna().any().any():
        return None

    m = P.to_numpy(dtype=float)

    if not np.isfinite(m).all():
        return None

    alcance = m > 0
    np.fill_diagonal(alcance, True)

    for k in range(3):
        for i in range(3):
            for j in range(3):
                alcance[i, j] = (
                    alcance[i, j]
                    or (
                        alcance[i, k]
                        and alcance[k, j]
                    )
                )

    if not alcance.all():
        return None

    A = np.vstack([
        m.T - np.eye(3),
        np.ones(3)
    ])

    b = np.array([0, 0, 0, 1])

    pi = np.linalg.lstsq(
        A, b, rcond=None
    )[0]

    pi = np.clip(pi, 0, None)

    if pi.sum() <= 0:
        return None

    return pi / pi.sum()


def permanencia(P):
    resultados = []

    for classe in CLASSES:
        p = P.loc[classe, classe]

        if pd.isna(p):
            resultados.append(np.nan)
        elif p >= 1 - 1e-12:
            resultados.append(np.inf)
        else:
            resultados.append(1 / (1 - p))

    return resultados


def identificar_estacao(mes):
    if mes in (12, 1, 2):
        return "Verão"
    if mes in (3, 4, 5):
        return "Outono"
    if mes in (6, 7, 8):
        return "Inverno"
    return "Primavera"


def markov_sazonal(tabela, ano):
    d = tabela[
        tabela["ano"] == ano
    ].copy()

    d["estacao"] = d["mes"].apply(
        identificar_estacao
    )

    resultados = []

    for estacao in ESTACOES:
        parte = d[
            d["estacao"] == estacao
        ].copy()

        N, P, n = matriz_markov(parte)

        pi = estacionaria(P)
        tempos = permanencia(P)

        if pi is None:
            pi = [np.nan] * 3

        resultados.append({
            "Ano": ano,
            "Estação": estacao,
            "π Fraco": pi[0],
            "π Moderado": pi[1],
            "π Forte": pi[2],
            "Dias Fraco": tempos[0],
            "Dias Moderado": tempos[1],
            "Dias Forte": tempos[2],
            "Total de dias": int(
                parte["vento"].notna().sum()
            ),
            "Transições válidas": n
        })

    return pd.DataFrame(resultados)


def formatar_sazonal(df):
    tabela = df.copy()

    for c in [
        "π Fraco", "π Moderado", "π Forte"
    ]:
        tabela[c] = tabela[c].apply(
            lambda x: numero(x, 3)
        )

    for c in [
        "Dias Fraco", "Dias Moderado",
        "Dias Forte"
    ]:
        def formatar(x):
            if pd.isna(x):
                return "—"
            if np.isinf(x):
                return "∞"
            return f"{numero(x)} ({round(x)})"

        tabela[c] = tabela[c].apply(formatar)

    return tabela


def episodios_vento(tabela):
    d = tabela.sort_values("data").copy()

    anterior = d["classe"].shift()
    mudou = (
        d["classe"].ne(anterior).fillna(True)
        | d["classe"].isna()
        | anterior.isna()
    )

    d["grupo"] = mudou.cumsum()

    return (
        d.dropna(subset=["classe"])
        .groupby(
            ["classe", "grupo"],
            as_index=False
        )
        .agg(dias=("data", "count"))
    )


def mapa_calor_markov(P):
    z = P.to_numpy(dtype=float)

    textos = [
        [
            f"{v:.1%}" if pd.notna(v) else "—"
            for v in linha
        ]
        for linha in z
    ]

    fig = go.Figure(
        go.Heatmap(
            z=z,
            x=CLASSES,
            y=CLASSES,
            zmin=0,
            zmax=1,
            colorscale=[
                [0, "#FFF9ED"],
                [0.5, "#F3CB64"],
                [1, "#865018"]
            ],
            text=textos,
            texttemplate="%{text}",
            textfont=dict(color="#111111"),
            hovertemplate=(
                "%{y} → %{x}<br>"
                "Probabilidade: %{z:.1%}"
                "<extra></extra>"
            )
        )
    )

    fig.update_layout(
        title="Matriz de transição geral",
        xaxis_title="Próximo dia",
        yaxis_title="Dia atual"
    )

    return fig


# ============================================================
# 8. TESTES ESTATÍSTICOS
# ============================================================

def testes_vento(serie):
    s = serie.dropna().astype(float)

    registros = []

    if len(s) < 36:
        return pd.DataFrame()

    try:
        a = adfuller(s, autolag="AIC")

        registros.append({
            "Teste": "ADF",
            "Estatística": a[0],
            "Valor-p": a[1],
            "Hipótese nula": "Raiz unitária"
        })
    except Exception:
        pass

    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            k = kpss(
                s,
                regression="c",
                nlags="auto"
            )

        registros.append({
            "Teste": "KPSS",
            "Estatística": k[0],
            "Valor-p": k[1],
            "Hipótese nula": "Estacionariedade"
        })
    except Exception:
        pass

    try:
        lb = acorr_ljungbox(
            s - s.mean(),
            lags=[12],
            return_df=True
        )

        registros.append({
            "Teste": "Ljung-Box",
            "Estatística": lb["lb_stat"].iloc[-1],
            "Valor-p": lb["lb_pvalue"].iloc[-1],
            "Hipótese nula": "Sem autocorrelação"
        })
    except Exception:
        pass

    anual = s.resample("YS").mean().dropna()

    if len(anual) >= 5:
        tau, p = stats.kendalltau(
            np.arange(len(anual)),
            anual.values
        )

        registros.append({
            "Teste": "Kendall anual",
            "Estatística": tau,
            "Valor-p": p,
            "Hipótese nula": "Sem associação monotônica"
        })

    return pd.DataFrame(registros)


# ============================================================
# 9. MAPAS E NAVEGAÇÃO
# ============================================================

def listar_png(pasta):
    if not pasta.is_dir():
        return []

    return sorted([
        p for p in pasta.iterdir()
        if p.suffix.lower() == ".png"
    ])


GALERIAS = {
    nome: listar_png(pasta)
    for nome, pasta in PASTAS.items()
}


def titulo_mapa(arquivo):
    anos = re.findall(
        r"(?:19|20)\d{2}",
        arquivo.stem
    )

    if len(anos) >= 2:
        return (
            f"Velocidade do vento | "
            f"{anos[0]}–{anos[1]}"
        )

    return arquivo.stem.replace("_", " ").title()


if "mapa_tela" not in st.session_state:
    st.session_state.mapa_tela = "menu"

if "mapa_categoria" not in st.session_state:
    st.session_state.mapa_categoria = None

if "mapa_imagem" not in st.session_state:
    st.session_state.mapa_imagem = None


def entrar_categoria(categoria):
    st.session_state.mapa_categoria = categoria
    st.session_state.mapa_tela = "galeria"


def entrar_imagem(caminho):
    st.session_state.mapa_imagem = str(caminho)
    st.session_state.mapa_tela = "imagem"


def voltar_galeria():
    st.session_state.mapa_tela = "galeria"


def voltar_menu():
    st.session_state.mapa_tela = "menu"


def galeria_mapas(arquivos):
    if not arquivos:
        st.warning("Nenhum PNG encontrado.")
        return

    for inicio in range(0, len(arquivos), 3):
        colunas = st.columns(3)

        for coluna, arquivo in zip(
            colunas,
            arquivos[inicio:inicio + 3]
        ):
            with coluna:
                with st.container(border=True):
                    st.markdown(
                        f"#### {titulo_mapa(arquivo)}"
                    )
                    st.image(
                        str(arquivo),
                        use_container_width=True
                    )
                    st.button(
                        "Visualizar mapa →",
                        key=f"abrir_{arquivo}",
                        on_click=entrar_imagem,
                        args=(arquivo,),
                        use_container_width=True
                    )


# ============================================================
# 10. BARRA LATERAL
# ============================================================

with st.sidebar:
    st.title("🌬️ Observatório dos Ventos")

    st.divider()
    st.subheader("Área de estudo")
    st.write("Paraíba e Pernambuco")
    st.write("1991–2025")

    st.divider()
    st.subheader("Base de dados")

    if dados is not None:
        st.success("Excel conectado")
        st.write(f"Registros: {len(dados):,}")
        st.write(f"Abas Excel: {len(abas_excel)}")
        st.write(
            f"Início: {dados['data'].min():%d/%m/%Y}"
        )
        st.write(
            f"Fim: {dados['data'].max():%d/%m/%Y}"
        )
    else:
        st.warning("Excel não encontrado.")
        if erro_excel:
            st.caption(erro_excel)

    st.divider()
    st.subheader("Cartografia")

    for nome, arquivos in GALERIAS.items():
        st.write(f"{nome}: {len(arquivos)}")

    st.divider()
    st.subheader("Autores")
    st.write("Denilson")
    st.write("Jader")
    st.write("Silvio")


# ============================================================
# 11. CABEÇALHO NATIVO
# ============================================================

with st.container(border=True):
    st.caption(
        "MONITORAMENTO AMBIENTAL • PARAÍBA E PERNAMBUCO"
    )
    st.title("🌬️ Observatório dos Ventos")
    st.write(
        "Dinâmica espaço-temporal da velocidade "
        "do vento | 1991–2025"
    )
    st.progress(100)

st.write("")


# ============================================================
# 12. ABAS
# ============================================================

abas = st.tabs([
    "Autores",
    "Documentação",
    "Mapas",
    "Descritiva",
    "Séries temporais",
    "Anomalias",
    "Cadeia de Markov",
    "Testes",
    "Conclusão"
])


# ============================================================
# ABA 1 — AUTORES
# ============================================================


# ============================================================
# ABA 1 — AUTORES COMPLETOS
# ============================================================


# ============================================================
# ABA 1 — AUTORES COMPLETOS
# ============================================================

with abas[0]:

    st.header("Equipe Científica")

    with st.container(border=True):

        st.caption(
            "PESQUISA • ESTATÍSTICA • AMBIENTE"
        )

        st.subheader("Pesquisadores responsáveis")

        st.write(
            "O Observatório dos Ventos reúne "
            "pesquisadores vinculados à Universidade "
            "Federal Rural de Pernambuco (UFRPE) e à "
            "Universidade Estadual da Paraíba (UEPB), "
            "integrando métodos estatísticos, "
            "séries temporais, cartografia ambiental "
            "e cadeias de Markov."
        )

    # ========================================================
    # CADASTRO COMPLETO DOS AUTORES
    # ========================================================

    autores = [
        {
            "nome": "Denilson de Oliveira Silva",
            "iniciais": "DS",
            "instituicao": "UFRPE",
            "vinculo": (
                "Programa de Pós-Graduação em "
                "Biometria e Estatística Aplicada"
            ),
            "localizacao": "Recife, Pernambuco, Brasil",
            "email": "denilson.oliveira@ufrpe.br",
            "orcid": "0009-0000-4031-7772",
            "lattes": None
        },
        {
            "nome": "Jader da Silva Jale",
            "iniciais": "JJ",
            "instituicao": "UFRPE",
            "vinculo": (
                "Programa de Pós-Graduação em "
                "Biometria e Estatística Aplicada"
            ),
            "localizacao": "Recife, Pernambuco, Brasil",
            "email": "jader.silvajale@ufrpe.br",
            "orcid": "0000-0001-7414-1154",
            "lattes": None
        },
        {
            "nome": "Silvio Fernando Alves Xavier",
            "iniciais": "SX",
            "instituicao": "UEPB",
            "vinculo": "Departamento de Estatística",
            "localizacao": "Recife, Pernambuco, Brasil",
            "email": "silvio@servidor.uepb.edu.br",
            "orcid": "0000-0002-4832-0711",
            "lattes": None
        },
        {
    "nome": "Débora  de Sousa Cordeiro",
    "iniciais": "DC",
    "instituicao": "UFRPE",
    "vinculo": (
        "Programa de Pós-Graduação em "
        "Biometria e Estatística Aplicada"
    ),
    "localizacao": "Recife, Pernambuco, Brasil",
    "email": "debora.scordeiro@ufrpe.br",
    "orcid": "0009-0006-1410-4718",
    "lattes": "http://lattes.cnpq.br/1980673256363690"
}
    ]

    # ========================================================
    # CARTÕES DOS PESQUISADORES — 2 COLUNAS
    # ========================================================

    st.write("")

    colunas = st.columns(2, gap="medium")

    for indice, autor in enumerate(autores):

        with colunas[indice % 2]:

            with st.container(border=True):

                st.caption(
                    "PESQUISADOR • " + autor["instituicao"]
                )

                st.title(autor["iniciais"])

                st.subheader(autor["nome"])

                st.markdown(
                    f"🏛️ **Instituição:** "
                    f"{autor['instituicao']}"
                )

                st.markdown(
                    f"🎓 **Vínculo acadêmico:** "
                    f"{autor['vinculo']}"
                )

                st.markdown(
                    f"📍 **Localização:** "
                    f"{autor['localizacao']}"
                )

                st.divider()

                # ============================================
                # CONTATO ACADÊMICO
                # ============================================

                st.markdown("#### 📬 Contato acadêmico")

                st.markdown(
                    f"**E-mail:** {autor['email']}"
                )

                st.link_button(
                    "Contato por e-mail ↗",
                    "mailto:" + autor["email"],
                    use_container_width=True
                )

                st.divider()

                # ============================================
                # ORCID
                # ============================================

                st.markdown(
                    "#### 🔗 Identificadores acadêmicos"
                )

                if autor.get("orcid"):

                    st.markdown(
                        f"**ORCID:** {autor['orcid']}"
                    )

                    st.link_button(
                        "Identificador ORCID ↗",
                        "https://orcid.org/" + autor["orcid"],
                        use_container_width=True
                    )

                else:

                    st.caption(
                        "ORCID não informado."
                    )

                # ============================================
                # CURRÍCULO LATTES
                # ============================================

                if autor.get("lattes"):

                    st.link_button(
                        "Currículo Lattes ↗",
                        autor["lattes"],
                        use_container_width=True
                    )

                else:

                    st.caption(
                        "Currículo Lattes não informado."
                    )

    # ========================================================
    # INSTITUIÇÕES PARTICIPANTES
    # ========================================================

    st.divider()

    st.subheader("Instituições participantes")

    c1, c2 = st.columns(2, gap="medium")

    with c1:

        with st.container(border=True):

            st.markdown(
                "### Universidade Federal Rural "
                "de Pernambuco (UFRPE)"
            )

            st.write(
                "Programa de Pós-Graduação em "
                "Biometria e Estatística Aplicada"
            )

            st.write(
                "📍 Recife, Pernambuco, Brasil"
            )

    with c2:

        with st.container(border=True):

            st.markdown(
                "### Universidade Estadual "
                "da Paraíba (UEPB)"
            )

            st.write(
                "Departamento de Estatística"
            )

            st.write(
                "📍 Campina Grande, Paraíba, Brasil"
            )

    # ========================================================
    # INTEGRAÇÃO CIENTÍFICA E TECNOLÓGICA
    # ========================================================

    st.divider()

    with st.container(border=True):

        st.subheader(
            "Integração científica e tecnológica"
        )

        st.write(
            "O Observatório dos Ventos integra "
            "pesquisadores da Universidade Federal "
            "Rural de Pernambuco (UFRPE) e da "
            "Universidade Estadual da Paraíba (UEPB), "
            "promovendo análises estatísticas e "
            "ambientais da velocidade do vento "
            "nos estados da Paraíba e Pernambuco."
        )

        st.markdown(
            "**Equipe científica:** "
            "Denilson de Oliveira Silva; "
            "Jader da Silva Jale; "
            "Silvio Fernando Alves Xavier; "
            "Débora Cordeiro de Sousa."
        )

# ============================================================
# ABA 2 — DOCUMENTAÇÃO
# ============================================================

with abas[1]:
    st.header("Documentação Científica")

    st.subheader("Resumo")
    st.write(
        "O Observatório dos Ventos é uma aplicação "
        "interativa dedicada à análise espacial e "
        "temporal da velocidade do vento nos estados "
        "da Paraíba e de Pernambuco entre 1991 e 2025. "
        "Integra registros diários, mapas temáticos, "
        "análises estatísticas e cadeias de Markov."
    )

    st.subheader("Materiais")
    st.write(
        "A base utiliza registros diários de velocidade "
        "do vento em metros por segundo, organizados "
        "por data e estado em uma planilha Excel. "
        "Os mapas foram preparados previamente no R "
        "e incorporados à plataforma como imagens PNG."
    )

    st.subheader("Métodos")
    st.write(
        "A análise contempla estatística descritiva, "
        "médias móveis, suavização exponencial, "
        "decomposição STL, anomalias e testes "
        "estatísticos. A cadeia de Markov utiliza "
        "três classes definidas pelos quartis "
        "da distribuição histórica de referência."
    )

    st.latex(
        r"\overline{V}=\frac{1}{n}"
        r"\sum_{t=1}^{n}V_t"
    )

    st.latex(
        r"MM_t(k)=\frac{1}{k}"
        r"\sum_{j=0}^{k-1}V_{t-j}"
    )

    st.latex(
        r"A_{m,t}=V_{m,t}-\overline{V}_m"
    )

    st.write(
        "A classificação adota vento fraco até Q1, "
        "moderado entre Q1 e Q3 e forte acima de Q3. "
        "Q2 representa a mediana da distribuição."
    )

    st.latex(
        r"F:V\leq Q_1,\quad"
        r"M:Q_1<V\leq Q_3,\quad"
        r"FO:V>Q_3"
    )

    st.latex(
        r"P_{ij}=\frac{N_{ij}}{\sum_jN_{ij}}"
    )

    st.latex(
        r"\boldsymbol{\pi}P=\boldsymbol{\pi}"
    )

    st.latex(
        r"E(D_i)=\frac{1}{1-P_{ii}}"
    )

    st.write(
        "A análise sazonal estima separadamente "
        "as matrizes de transição para as quatro "
        "estações de cada ano. As probabilidades "
        "estacionárias são apresentadas somente "
        "quando a cadeia possui uma distribuição "
        "de equilíbrio única identificável."
    )


# ============================================================
# ABA 3 — MAPAS
# ============================================================

with abas[2]:
    st.header("Cartografia da Velocidade do Vento")

    tela = st.session_state.mapa_tela
    categoria = st.session_state.mapa_categoria

    if tela == "menu":
        colunas = st.columns(3)

        for coluna, (nome, arquivos) in zip(
            colunas, GALERIAS.items()
        ):
            with coluna:
                with st.container(border=True):
                    st.subheader(nome)
                    st.metric(
                        "Arquivos disponíveis",
                        len(arquivos)
                    )
                    st.button(
                        "Visualizar →",
                        key=f"categoria_{nome}",
                        use_container_width=True,
                        on_click=entrar_categoria,
                        args=(nome,)
                    )

    elif tela == "galeria":
        st.button(
            "← Voltar ao menu",
            on_click=voltar_menu,
            key="voltar_menu"
        )
        st.subheader(categoria)
        galeria_mapas(GALERIAS[categoria])

    elif tela == "imagem":
        c1, c2 = st.columns(2)

        with c1:
            st.button(
                "← Voltar à galeria",
                key="voltar_galeria",
                on_click=voltar_galeria,
                use_container_width=True
            )

        with c2:
            st.button(
                "⌂ Menu de mapas",
                key="menu_mapas",
                on_click=voltar_menu,
                use_container_width=True
            )

        caminho = st.session_state.mapa_imagem
        arquivo = Path(caminho) if caminho else None

        if arquivo and arquivo.is_file():
            st.subheader(titulo_mapa(arquivo))
            st.image(
                str(arquivo),
                use_container_width=True
            )

            arquivos = GALERIAS[categoria]

            if arquivo in arquivos:
                pos = arquivos.index(arquivo)
                a, b, c = st.columns([1, 2, 1])

                with a:
                    if pos > 0:
                        st.button(
                            "← Anterior",
                            key="anterior",
                            on_click=entrar_imagem,
                            args=(arquivos[pos - 1],)
                        )

                with b:
                    st.caption(
                        f"Mapa {pos + 1} de {len(arquivos)}"
                    )

                with c:
                    if pos < len(arquivos) - 1:
                        st.button(
                            "Próximo →",
                            key="proximo",
                            on_click=entrar_imagem,
                            args=(arquivos[pos + 1],)
                        )


# ============================================================
# ABA 4 — DESCRITIVA
# ============================================================

with abas[3]:
    st.header("Análise Descritiva")

    if dados is None:
        st.info("Aguardando dados.")
    else:
        intervalo = st.slider(
            "Intervalo de anos",
            int(dados["ano"].min()),
            int(dados["ano"].max()),
            (
                int(dados["ano"].min()),
                int(dados["ano"].max())
            )
        )

        df = dados[
            dados["ano"].between(*intervalo)
        ].dropna(subset=["vento"]).copy()

        if df.empty:
            st.warning("Sem dados no intervalo.")
        else:
            c1, c2, c3, c4 = st.columns(4)

            c1.metric(
                "Velocidade média",
                numero(df["vento"].mean()) + " m/s"
            )
            c2.metric(
                "Desvio-padrão",
                numero(df["vento"].std()) + " m/s"
            )
            c3.metric(
                "Mínimo",
                numero(df["vento"].min()) + " m/s"
            )
            c4.metric(
                "Máximo",
                numero(df["vento"].max()) + " m/s"
            )

            a, b = st.columns(2)

            with a:
                fig = px.histogram(
                    df,
                    x="vento",
                    color="estado",
                    nbins=45,
                    barmode="overlay",
                    opacity=0.65,
                    color_discrete_map=CORES_ESTADOS,
                    title="Distribuição das velocidades"
                )
                grafico(fig)

            with b:
                fig = px.box(
                    df,
                    x="estado",
                    y="vento",
                    color="estado",
                    color_discrete_map=CORES_ESTADOS,
                    title="Boxplots por estado"
                )
                grafico(fig)

            c, d = st.columns(2)

            with c:
                fig = px.violin(
                    df,
                    x="estado",
                    y="vento",
                    color="estado",
                    box=True,
                    color_discrete_map=CORES_ESTADOS,
                    title="Distribuição e quartis"
                )
                grafico(fig)

            with d:
                mensal = df.groupby(
                    ["estado", "mes"],
                    as_index=False
                )["vento"].mean()

                fig = px.line(
                    mensal,
                    x="mes",
                    y="vento",
                    color="estado",
                    markers=True,
                    color_discrete_map=CORES_ESTADOS,
                    title="Ciclo médio mensal"
                )
                grafico(fig)

            e, f = st.columns(2)

            with e:
                resumo_periodos = df.groupby(
                    ["estado", "periodo"],
                    as_index=False
                )["vento"].mean()

                fig = px.bar(
                    resumo_periodos,
                    x="periodo",
                    y="vento",
                    color="estado",
                    barmode="group",
                    category_orders={
                        "periodo": PERIODOS
                    },
                    color_discrete_map=CORES_ESTADOS,
                    title="Velocidade por período"
                )
                grafico(fig)

            with f:
                fig = px.ecdf(
                    df,
                    x="vento",
                    color="estado",
                    color_discrete_map=CORES_ESTADOS,
                    title="Distribuição acumulada"
                )
                grafico(fig)

            resumo = df.groupby(
                "estado"
            )["vento"].agg([
                "count", "mean", "median",
                "std", "min", "max"
            ]).reset_index()

            st.subheader("Resumo estatístico")
            st.dataframe(
                resumo,
                use_container_width=True,
                hide_index=True
            )
            baixar_csv(
                resumo,
                "descritiva_vento.csv",
                "baixar_descritiva"
            )


# ============================================================
# ABA 5 — SÉRIES TEMPORAIS
# ============================================================

with abas[4]:
    st.header("Séries Temporais")

    if dados is None:
        st.info("Aguardando dados.")
    else:
        estado = st.selectbox(
            "Estado", ESTADOS,
            key="estado_series"
        )

        diaria = serie_estado(dados, estado)
        mensal = serie_mensal(dados, estado)

        tecnica = st.radio(
            "Técnica",
            [
                "Série bruta",
                "Médias móveis",
                "Suavização exponencial",
                "Decomposição STL",
                "Autocorrelação"
            ],
            horizontal=True
        )

        if tecnica == "Série bruta":
            fig = go.Figure()

            fig.add_trace(
                go.Scattergl(
                    x=diaria.index,
                    y=diaria.values,
                    mode="lines",
                    name=estado,
                    line=dict(color=DOURADO, width=1)
                )
            )

            fig.update_layout(
                title="Velocidade diária do vento",
                yaxis_title="Velocidade (m/s)"
            )
            grafico(fig, 500)

            anual = diaria.resample("YS").mean()

            fig = px.line(
                x=anual.index,
                y=anual.values,
                markers=True,
                title="Médias anuais"
            )
            fig.update_traces(line_color=MARROM)
            grafico(fig)

        elif tecnica == "Médias móveis":
            janelas = st.multiselect(
                "Janelas em dias",
                [7, 15, 30, 60, 90, 180, 365],
                default=[30, 90, 365]
            )

            fig = go.Figure()

            fig.add_trace(
                go.Scattergl(
                    x=diaria.index,
                    y=diaria.values,
                    name="Original",
                    line=dict(
                        color="#D9D4CA",
                        width=1
                    )
                )
            )

            for i, janela in enumerate(janelas):
                media = diaria.rolling(
                    janela,
                    min_periods=max(1, janela // 2)
                ).mean()

                fig.add_trace(
                    go.Scatter(
                        x=media.index,
                        y=media.values,
                        name=f"MM {janela} dias",
                        line=dict(
                            color=CORES_SERIES[i],
                            width=2.5
                        )
                    )
                )

            fig.update_layout(
                title="Médias móveis",
                legend=dict(
                    orientation="h",
                    y=1.15,
                    x=0,
                    font=dict(color="#111111")
                )
            )
            grafico(fig, 510)

        elif tecnica == "Suavização exponencial":
            alfa = st.slider(
                "Parâmetro alfa",
                0.05, 0.95, 0.25, 0.05
            )

            suavizada = mensal.ewm(
                alpha=alfa,
                adjust=False
            ).mean()

            fig = go.Figure()

            fig.add_trace(
                go.Scatter(
                    x=mensal.index,
                    y=mensal.values,
                    name="Observado",
                    line=dict(color=DOURADO)
                )
            )

            fig.add_trace(
                go.Scatter(
                    x=suavizada.index,
                    y=suavizada.values,
                    name="Suavizado",
                    line=dict(
                        color=MARROM, width=3
                    )
                )
            )

            fig.update_layout(
                title="Suavização exponencial"
            )
            grafico(fig, 500)

        elif tecnica == "Decomposição STL":
            s = mensal.asfreq("MS")

            if (
                len(s.dropna()) < 36
                or s.isna().mean() > 0.05
            ):
                st.warning(
                    "Série insuficiente ou com "
                    "muitas lacunas."
                )
            else:
                s = s.interpolate(
                    method="time",
                    limit=2,
                    limit_direction="both"
                )

                if s.isna().any():
                    st.warning(
                        "Lacunas não resolvidas."
                    )
                else:
                    try:
                        ajuste = STL(
                            s,
                            period=12,
                            robust=True
                        ).fit()

                        componentes = [
                            ("Observada", s),
                            ("Tendência", ajuste.trend),
                            ("Sazonalidade", ajuste.seasonal),
                            ("Resíduos", ajuste.resid)
                        ]

                        for i in range(0, 4, 2):
                            colunas = st.columns(2)

                            for col, (nome, comp) in zip(
                                colunas,
                                componentes[i:i + 2]
                            ):
                                with col:
                                    fig = px.line(
                                        x=comp.index,
                                        y=comp.values,
                                        title=nome
                                    )
                                    fig.update_traces(
                                        line_color=(
                                            DOURADO if i == 0
                                            else MARROM
                                        )
                                    )
                                    grafico(fig)

                    except Exception as erro:
                        st.warning(str(erro))

        elif tecnica == "Autocorrelação":
            s = mensal.dropna()

            if len(s) < 30:
                st.warning("Série insuficiente.")
            else:
                valores = acf(
                    s,
                    nlags=min(36, len(s) // 3),
                    fft=True
                )

                fig = px.bar(
                    x=np.arange(len(valores)),
                    y=valores,
                    title="Autocorrelação mensal",
                    color_discrete_sequence=[DOURADO]
                )
                grafico(fig)


# ============================================================
# ABA 6 — ANOMALIAS
# ============================================================

with abas[5]:
    st.header("Anomalias do Vento")

    if dados is None:
        st.info("Aguardando dados.")
    else:
        estado = st.selectbox(
            "Estado", ESTADOS,
            key="estado_anomalias"
        )

        tabela = anomalias_mensais(
            dados, estado
        )

        escala = [
            "#70451C", "#EAB744", "#FFFFFF",
            "#D89323", "#865018"
        ]

        fig = px.bar(
            tabela,
            x="data",
            y="anomalia",
            color="anomalia",
            color_continuous_scale=escala,
            color_continuous_midpoint=0,
            title="Anomalias mensais"
        )

        fig.add_hline(
            y=0,
            line_color=ESCURO
        )
        grafico(fig, 480)

        anual = tabela.groupby(
            "ano",
            as_index=False
        )["anomalia"].mean()

        fig = px.bar(
            anual,
            x="ano",
            y="anomalia",
            color="anomalia",
            color_continuous_scale=escala,
            color_continuous_midpoint=0,
            title="Anomalias médias anuais"
        )
        grafico(fig)

        baixar_csv(
            tabela,
            "anomalias_vento.csv",
            "baixar_anomalias"
        )


# ============================================================
# ABA 7 — CADEIA DE MARKOV
# ============================================================

with abas[6]:
    st.header("Cadeias de Markov do Vento")

    if dados is None:
        st.info("Aguardando dados.")
    else:
        estado = st.selectbox(
            "Estado",
            ESTADOS,
            key="estado_markov"
        )

        referencia = st.radio(
            "Referência dos quartis",
            [
                "Quartis do próprio estado",
                "Quartis conjuntos"
            ],
            horizontal=True
        )

        ref = (
            dados[dados["estado"] == estado]
            if referencia == "Quartis do próprio estado"
            else dados
        )

        try:
            quartis = quartis_vento(ref)
            q1, q2, q3 = quartis

            c1, c2, c3 = st.columns(3)

            c1.metric(
                "Q1 — Fraco",
                numero(q1) + " m/s"
            )
            c2.metric(
                "Q2 — Mediana",
                numero(q2) + " m/s"
            )
            c3.metric(
                "Q3 — Forte",
                numero(q3) + " m/s"
            )

            tabela = preparar_markov(
                dados, estado, quartis
            )

            N, P, n = matriz_markov(tabela)

            st.subheader("Markov geral — 1991–2025")

            ca, cb = st.columns(2)

            with ca:
                freq = (
                    tabela["classe"]
                    .value_counts()
                    .reindex(CLASSES, fill_value=0)
                    .reset_index()
                )
                freq.columns = [
                    "Classe", "Frequência"
                ]

                fig = px.bar(
                    freq,
                    x="Classe",
                    y="Frequência",
                    color="Classe",
                    color_discrete_map=CORES_MARKOV,
                    title="Frequência por categoria"
                )
                grafico(fig)

            with cb:
                grafico(mapa_calor_markov(P))

            st.caption(
                f"{n:,} transições diárias válidas."
            )

            st.subheader("Matriz de probabilidades")

            st.dataframe(
                P.style.format(
                    "{:.3f}",
                    na_rep="—"
                ),
                use_container_width=True
            )

            st.subheader(
                "Probabilidades estacionárias"
            )

            pi = estacionaria(P)

            if pi is None:
                st.info(
                    "Distribuição estacionária única "
                    "não identificável."
                )
            else:
                resultado_pi = pd.DataFrame({
                    "Classe": CLASSES,
                    "Probabilidade": pi
                })

                fig = px.bar(
                    resultado_pi,
                    x="Classe",
                    y="Probabilidade",
                    color="Classe",
                    color_discrete_map=CORES_MARKOV,
                    title="Distribuição estacionária"
                )
                fig.update_layout(
                    yaxis_range=[0, 1]
                )
                grafico(fig)

            st.subheader("Permanência esperada")

            tempos = permanencia(P)
            cols = st.columns(3)

            for col, classe, tempo in zip(
                cols, CLASSES, tempos
            ):
                with col:
                    if np.isinf(tempo):
                        valor = "∞"
                    elif pd.notna(tempo):
                        valor = numero(tempo) + " dias"
                    else:
                        valor = "—"

                    st.metric(
                        classe, valor
                    )

            episodios = episodios_vento(tabela)

            if not episodios.empty:
                fig = px.box(
                    episodios,
                    x="classe",
                    y="dias",
                    color="classe",
                    color_discrete_map=CORES_MARKOV,
                    title="Duração observada dos episódios"
                )
                grafico(fig)

            baixar_csv(
                P.reset_index(),
                "markov_geral.csv",
                "baixar_markov_geral"
            )

            # --------------------------------------------
            # MARKOV ANUAL E SAZONAL
            # --------------------------------------------

            st.divider()

            st.header(
                "Probabilidades estacionárias "
                "por ano e estação"
            )

            st.write(
                "Selecione um ano para examinar "
                "a distribuição estacionária dos "
                "ventos fraco, moderado e forte "
                "e os tempos esperados de permanência "
                "em cada estação."
            )

            anos = sorted(
                tabela.loc[
                    tabela["vento"].notna(),
                    "ano"
                ].dropna().astype(int).unique()
            )

            if len(anos) == 0:
                st.warning("Nenhum ano disponível.")
            else:
                ano = st.selectbox(
                    "Ano de análise",
                    anos,
                    index=len(anos) - 1,
                    key="ano_markov"
                )

                sazonal = markov_sazonal(
                    tabela, ano
                )

                st.subheader(
                    f"Tabela sazonal — {estado} | {ano}"
                )

                st.dataframe(
                    formatar_sazonal(sazonal),
                    use_container_width=True,
                    hide_index=True
                )

                baixar_csv(
                    sazonal,
                    f"markov_{estado}_{ano}.csv",
                    "baixar_markov_sazonal"
                )

                probabilidades = sazonal.melt(
                    id_vars=["Estação"],
                    value_vars=[
                        "π Fraco",
                        "π Moderado",
                        "π Forte"
                    ],
                    var_name="Categoria",
                    value_name="Probabilidade"
                )

                probabilidades["Categoria"] = (
                    probabilidades["Categoria"]
                    .str.replace(
                        "π ", "", regex=False
                    )
                )

                tempos_sazonais = sazonal.melt(
                    id_vars=["Estação"],
                    value_vars=[
                        "Dias Fraco",
                        "Dias Moderado",
                        "Dias Forte"
                    ],
                    var_name="Categoria",
                    value_name="Dias"
                )

                tempos_sazonais["Categoria"] = (
                    tempos_sazonais["Categoria"]
                    .str.replace(
                        "Dias ", "", regex=False
                    )
                )

                tempos_sazonais["Dias"] = (
                    tempos_sazonais["Dias"]
                    .replace(
                        [np.inf, -np.inf],
                        np.nan
                    )
                )

                ca, cb = st.columns(2)

                with ca:
                    fig = px.bar(
                        probabilidades,
                        x="Estação",
                        y="Probabilidade",
                        color="Categoria",
                        barmode="group",
                        category_orders={
                            "Estação": ESTACOES,
                            "Categoria": CLASSES
                        },
                        color_discrete_map=CORES_MARKOV,
                        title="Probabilidades estacionárias"
                    )
                    fig.update_layout(
                        yaxis_range=[0, 1]
                    )
                    grafico(fig)

                with cb:
                    fig = px.bar(
                        tempos_sazonais,
                        x="Estação",
                        y="Dias",
                        color="Categoria",
                        barmode="group",
                        category_orders={
                            "Estação": ESTACOES,
                            "Categoria": CLASSES
                        },
                        color_discrete_map=CORES_MARKOV,
                        title="Permanência esperada"
                    )
                    grafico(fig)

                st.subheader(
                    "Cobertura das observações"
                )

                cobertura = sazonal[
                    [
                        "Estação",
                        "Total de dias",
                        "Transições válidas"
                    ]
                ]

                fig = px.bar(
                    cobertura,
                    x="Estação",
                    y=[
                        "Total de dias",
                        "Transições válidas"
                    ],
                    barmode="group",
                    color_discrete_sequence=[
                        DOURADO, MARROM
                    ],
                    title="Disponibilidade sazonal"
                )
                grafico(fig)

                if (
                    sazonal["Transições válidas"].min()
                    < 30
                ):
                    st.warning(
                        "Uma ou mais estações possuem "
                        "menos de 30 transições válidas. "
                        "Interprete as estimativas "
                        "com cautela."
                    )

                st.write(
                    "As probabilidades estacionárias "
                    "representam propriedades da "
                    "cadeia estimada. As durações "
                    "esperadas são calculadas com "
                    "1/(1−Pii). Estações com dados "
                    "insuficientes podem apresentar "
                    "resultados não identificáveis."
                )

        except Exception as erro:
            st.warning(
                f"Erro na análise de Markov: {erro}"
            )


# ============================================================
# ABA 8 — TESTES
# ============================================================

with abas[7]:
    st.header("Testes Estatísticos")

    if dados is None:
        st.info("Aguardando dados.")
    else:
        estado = st.selectbox(
            "Estado",
            ESTADOS,
            key="estado_testes"
        )

        mensal = serie_mensal(
            dados, estado
        )

        resultado = testes_vento(mensal)

        if resultado.empty:
            st.warning("Dados insuficientes.")
        else:
            visual = resultado.copy()

            visual["Estatística"] = (
                visual["Estatística"]
                .apply(
                    lambda x: numero(x, 4)
                )
            )

            visual["Valor-p"] = (
                visual["Valor-p"]
                .apply(valor_p)
            )

            st.dataframe(
                visual,
                use_container_width=True,
                hide_index=True
            )

            ca, cb = st.columns(2)

            with ca:
                fig = px.bar(
                    resultado,
                    x="Teste",
                    y="Valor-p",
                    color_discrete_sequence=[DOURADO],
                    title="Valores-p dos testes"
                )

                fig.add_hline(
                    y=0.05,
                    line_dash="dash",
                    line_color=MARROM
                )
                grafico(fig)

            with cb:
                diferencas = mensal.diff().dropna()

                fig = px.histogram(
                    x=diferencas,
                    nbins=40,
                    color_discrete_sequence=[MARROM],
                    title="Diferenças mensais"
                )
                grafico(fig)

            baixar_csv(
                resultado,
                "testes_vento.csv",
                "baixar_testes"
            )

            st.write(
                "Os testes ADF e KPSS investigam "
                "estacionariedade sob hipóteses "
                "nulas diferentes. Ljung–Box "
                "avalia autocorrelação temporal "
                "e Kendall examina associação "
                "monotônica das médias anuais. "
                "Os valores-p muito pequenos "
                "são apresentados como p < 0,01."
            )


# ============================================================
# ABA 9 — CONCLUSÃO
# ============================================================

with abas[8]:
    st.header("Conclusão")

    if dados is None:
        st.info("Aguardando a planilha Excel.")
    else:
        for estado in ESTADOS:
            d = dados[
                dados["estado"] == estado
            ].dropna(subset=["vento"])

            if d.empty:
                continue

            media = d["vento"].mean()

            periodos = (
                d.groupby("periodo")["vento"]
                .mean()
            )

            maior = periodos.idxmax()
            menor = periodos.idxmin()

            st.subheader(estado)

            a, b, c = st.columns(3)

            a.metric(
                "Velocidade média",
                numero(media) + " m/s"
            )

            b.metric(
                "Maior média por período",
                numero(periodos.max()) + " m/s"
            )

            c.metric(
                "Menor média por período",
                numero(periodos.min()) + " m/s"
            )

            st.write(
                f"Em {estado}, a velocidade média "
                f"dos registros disponíveis foi "
                f"de {numero(media)} m/s. O período "
                f"com maior média foi {maior}, "
                f"enquanto a menor média ocorreu "
                f"em {menor}. As análises temporais "
                f"e as cadeias de Markov permitem "
                f"examinar a variabilidade e a "
                f"persistência das categorias."
            )

        st.subheader("Síntese geral")

        st.write(
            "O Observatório dos Ventos integra "
            "cartografia, estatística descritiva, "
            "séries temporais e cadeias de Markov "
            "para caracterizar a dinâmica do vento "
            "na Paraíba e em Pernambuco entre "
            "1991 e 2025. As probabilidades "
            "estacionárias e os tempos de permanência "
            "permitem investigar padrões de "
            "persistência e alternância entre "
            "classes de intensidade. A interpretação "
            "deve considerar o tamanho amostral, "
            "a estabilidade das matrizes e as "
            "hipóteses dos métodos utilizados."
        )


# ============================================================
# RODAPÉ
# ============================================================

st.divider()

st.caption(
    "OBSERVATÓRIO DOS VENTOS | "
    "Paraíba e Pernambuco | 1991–2025"
)

st.caption(
    "Denilson • Jader • Silvio • Débora"
)
