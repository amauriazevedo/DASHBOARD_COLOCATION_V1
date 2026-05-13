
from pathlib import Path
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(
    page_title="CRONOGRAMA - PROJETO COLOCATION",
    page_icon="📊",
    layout="wide"
)


# =========================================================
# LOGIN
# =========================================================
USUARIO_CORRETO = "Ary"
SENHA_CORRETA = "Ary"

if "logado" not in st.session_state:
    st.session_state.logado = False

if not st.session_state.logado:

    st.markdown("""
    <style>
    .stApp {
        background: linear-gradient(135deg,#111827,#0F172A);
    }

    .login-box {
        background: white;
        padding: 45px;
        border-radius: 22px;
        max-width: 430px;
        margin: 100px auto;
        box-shadow: 0 15px 40px rgba(0,0,0,0.35);
    }

    .login-title {
        color: #111827;
        text-align: center;
        font-size: 36px;
        font-weight: 900;
        margin-bottom: 8px;
    }

    .login-sub {
        color: #6B7280;
        text-align: center;
        margin-bottom: 30px;
        font-size: 15px;
    }
    </style>
    """, unsafe_allow_html=True)

    st.markdown('<div class="login-box">', unsafe_allow_html=True)

    st.markdown('<div class="login-title">DASHBOARD COLOCATION V1</div>', unsafe_allow_html=True)
    st.markdown('<div class="login-sub">Acesso restrito ao sistema</div>', unsafe_allow_html=True)

    usuario = st.text_input("Usuário")
    senha = st.text_input("Senha", type="password")

    if st.button("Entrar", use_container_width=True):

        if usuario == USUARIO_CORRETO and senha == SENHA_CORRETA:
            st.session_state.logado = True
            st.rerun()
        else:
            st.error("Usuário ou senha inválidos.")

    st.markdown('</div>', unsafe_allow_html=True)

    st.stop()



st.markdown('''
<style>
.stApp {
    background: #F8FAFC;
    color: #111827;
}

header[data-testid="stHeader"] {
    background: #F8FAFC;
    height: 0rem;
}

.block-container {
    padding-top: 1rem;
    padding-left: 2rem;
    padding-right: 2rem;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #111827 0%, #0F172A 100%);
}

section[data-testid="stSidebar"] * {
    color: #FFFFFF !important;
}

h1 {
    color: #111827;
    font-size: 38px;
    font-weight: 900;
    letter-spacing: -1px;
}

h2, h3, h4 {
    color: #111827;
}

hr {
    border: none;
    border-top: 1px solid #E5E7EB;
}

.card {
    background: #FFFFFF;
    color: #111827;
    padding: 20px;
    border-radius: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 10px rgba(15,23,42,0.05);
    min-height: 105px;
}

.card-title {
    color: #64748B;
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
    margin-bottom: 14px;
}

.card-value {
    color: #111827;
    font-size: 23px;
    font-weight: 800;
}

.kpi-value {
    color: #2563EB;
    font-size: 32px;
    font-weight: 900;
}

.info-box {
    background: #FFFFFF;
    color: #111827;
    padding: 18px;
    border-radius: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 10px rgba(15,23,42,0.05);
    min-height: 105px;
}

.info-label {
    color: #64748B;
    font-size: 12px;
    font-weight: 800;
    text-transform: uppercase;
    margin-bottom: 12px;
}

.info-value {
    color: #111827;
    font-size: 16px;
    font-weight: 700;
}

.big-box {
    background: #FFFFFF;
    color: #111827;
    padding: 20px;
    border-radius: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 10px rgba(15,23,42,0.05);
    min-height: 210px;
}

.chart-box {
    background: #FFFFFF;
    padding: 10px;
    border-radius: 14px;
    border: 1px solid #E5E7EB;
    box-shadow: 0 2px 10px rgba(15,23,42,0.05);
}

div[data-testid="stSelectbox"] div {
    color: #111827 !important;
}

div[data-baseweb="select"] > div {
    background-color: #FFFFFF !important;
    color: #111827 !important;
    border-radius: 10px !important;
}

div[data-testid="stRadio"] label {
    color: #FFFFFF !important;
}

div[data-testid="stDataFrame"] {
    background: #FFFFFF;
}
</style>
''', unsafe_allow_html=True)


def normalizar_colunas(df):
    df.columns = (
        df.columns.astype(str)
        .str.replace("\\n", " ", regex=False)
        .str.replace("\\r", " ", regex=False)
        .str.strip()
    )
    return df


def achar_coluna(df, termos):
    for col in df.columns:
        txt = str(col).upper()
        if all(t.upper() in txt for t in termos):
            return col
    return None


def limpar_numero(serie):
    return pd.to_numeric(
        serie.astype(str)
        .str.replace("%", "", regex=False)
        .str.replace(",", ".", regex=False)
        .str.strip(),
        errors="coerce"
    )


def fmt_data(v):
    if pd.isna(v):
        return "-"
    try:
        return pd.to_datetime(v).strftime("%d/%m/%Y")
    except Exception:
        return str(v)


def fmt_valor(v):
    if pd.isna(v) or str(v).lower() == "nan":
        return "-"
    return str(v)


def carregar():
    caminho = Path("dados") / "Projeto Colocation.xlsx"
    if not caminho.exists():
        st.error("Planilha não encontrada. Coloque em: dados/Projeto Colocation.xlsx")
        st.stop()
    return pd.read_excel(caminho)


df = carregar()
df = normalizar_colunas(df).dropna(how="all")

col_id = achar_coluna(df, ["ID"])
col_titulo = achar_coluna(df, ["TÍTULO"])
col_desc = achar_coluna(df, ["DESCRIÇÃO"])
col_resp = achar_coluna(df, ["RESPONSÁVEL"])
col_percent = achar_coluna(df, ["CONCLUS"])
col_inicio_prog = achar_coluna(df, ["INÍCIO", "PROGRAMADO"])
col_termino_prog = achar_coluna(df, ["TÉRMINO", "PROGRAMADO"])
col_inicio_real = achar_coluna(df, ["INÍCIO", "REAL"])
col_termino_real = achar_coluna(df, ["TÉRMINO", "REAL"])
col_horas = achar_coluna(df, ["HORAS"])
col_duracao = achar_coluna(df, ["DURAÇÃO"])

if not col_titulo or not col_percent:
    st.error("Não encontrei as colunas obrigatórias: TÍTULO e CONCLUSÃO.")
    st.write(list(df.columns))
    st.stop()

df[col_percent] = limpar_numero(df[col_percent]).fillna(0)

if col_horas:
    df[col_horas] = limpar_numero(df[col_horas]).fillna(0)
else:
    df["TEMPO EM HORAS"] = 0
    col_horas = "TEMPO EM HORAS"

if col_duracao:
    df[col_duracao] = limpar_numero(df[col_duracao])

for c in [col_inicio_prog, col_termino_prog, col_inicio_real, col_termino_real]:
    if c:
        df[c] = pd.to_datetime(df[c], errors="coerce")

df = df[df[col_titulo].notna()].copy()
df[col_titulo] = df[col_titulo].astype(str).str.strip()

df["STATUS"] = df[col_percent].apply(lambda x: "Concluídas" if x >= 100 else "Não concluídas")
df["TAREFA_MENU"] = df.apply(
    lambda r: f"{r[col_id]} - {r[col_titulo]}" if col_id and pd.notna(r[col_id]) else r[col_titulo],
    axis=1
)

st.sidebar.markdown("## Menu")
visualizacao = st.sidebar.radio("Visualização", ["Visão geral", "Análise por tarefa"])
tarefa = st.sidebar.selectbox("Escolha uma tarefa", df["TAREFA_MENU"].tolist())

linha = df[df["TAREFA_MENU"] == tarefa].iloc[0]

st.title("DASHBOARD COLOCATION V1")

c1, c2, c3, c4 = st.columns(4)
with c1:
    st.markdown('<div class="card"><div class="card-title">TÍTULO DO PROJETO</div><div class="card-value">Projeto Colocation</div></div>', unsafe_allow_html=True)
with c2:
    st.markdown('<div class="card"><div class="card-title">NOME DA EMPRESA</div><div class="card-value">Imprensa Nacional</div></div>', unsafe_allow_html=True)
with c3:
    st.markdown('<div class="card"><div class="card-title">GERENTE DO PROJETO</div><div class="card-value">A definir</div></div>', unsafe_allow_html=True)
with c4:
    st.markdown('<div class="card"><div class="card-title">DATA DE INÍCIO PROJETADA</div><div class="card-value">20/05/2026</div></div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

total = len(df)
concluidas = int((df[col_percent] >= 100).sum())
nao_concluidas = total - concluidas
horas_totais = round(float(df[col_horas].sum()), 2)

k1, k2, k3, k4 = st.columns(4)
with k1:
    st.markdown(f'<div class="card"><div class="card-title">Total de tarefas</div><div class="kpi-value">{total}</div></div>', unsafe_allow_html=True)
with k2:
    st.markdown(f'<div class="card"><div class="card-title">Concluídas</div><div class="kpi-value">{concluidas}</div></div>', unsafe_allow_html=True)
with k3:
    st.markdown(f'<div class="card"><div class="card-title">Não concluídas</div><div class="kpi-value">{nao_concluidas}</div></div>', unsafe_allow_html=True)
with k4:
    st.markdown(f'<div class="card"><div class="card-title">Horas totais</div><div class="kpi-value">{horas_totais}</div></div>', unsafe_allow_html=True)

st.markdown("---")

st.subheader("📋 Dados da tarefa selecionada")

i1, i2, i3, i4 = st.columns(4)
with i1:
    st.markdown(f'<div class="info-box"><div class="info-label">TÍTULO</div><div class="info-value">{fmt_valor(linha.get(col_titulo))}</div></div>', unsafe_allow_html=True)
with i2:
    st.markdown(f'<div class="info-box"><div class="info-label">RESPONSÁVEL</div><div class="info-value">{fmt_valor(linha.get(col_resp)) if col_resp else "-"}</div></div>', unsafe_allow_html=True)
with i3:
    st.markdown(f'<div class="info-box"><div class="info-label">CONCLUSÃO</div><div class="info-value">{linha.get(col_percent, 0)}%</div></div>', unsafe_allow_html=True)
with i4:
    st.markdown(f'<div class="info-box"><div class="info-label">TEMPO EM HORAS</div><div class="info-value">{linha.get(col_horas, 0)}</div></div>', unsafe_allow_html=True)

d1, d2 = st.columns([2, 2])
with d1:
    desc = fmt_valor(linha.get(col_desc)) if col_desc else "-"
    st.markdown(f'<div class="big-box"><h4>Descrição</h4><p>{desc}</p></div>', unsafe_allow_html=True)

with d2:
    dur = fmt_valor(linha.get(col_duracao)) if col_duracao else "-"
    st.markdown(f'''
    <div class="big-box">
        <h4>Datas e Duração</h4>
        <p><b>Início programado:</b> {fmt_data(linha.get(col_inicio_prog)) if col_inicio_prog else "-"}</p>
        <p><b>Término programado:</b> {fmt_data(linha.get(col_termino_prog)) if col_termino_prog else "-"}</p>
        <p><b>Início real:</b> {fmt_data(linha.get(col_inicio_real)) if col_inicio_real else "-"}</p>
        <p><b>Término real:</b> {fmt_data(linha.get(col_termino_real)) if col_termino_real else "-"}</p>
        <p><b>Duração em dias:</b> {dur}</p>
    </div>
    ''', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

g1, g2 = st.columns(2)

df_pizza = pd.DataFrame({
    "Status": ["Concluídas", "Não concluídas"],
    "Quantidade": [concluidas, nao_concluidas],
})

fig_pizza = px.pie(
    df_pizza,
    names="Status",
    values="Quantidade",
    title="% DE CONCLUSÃO DAS TAREFAS (GERAL)",
    color="Status",
    color_discrete_map={"Concluídas": "#2563EB", "Não concluídas": "#EF4444"}
)
fig_pizza.update_traces(textinfo="percent")
fig_pizza.update_layout(paper_bgcolor="white", plot_bgcolor="white", font_color="#111827", legend_title_text="", height=460)

with g1:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.plotly_chart(fig_pizza, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

df_total_horas = pd.DataFrame({"Indicador": ["Total de horas"], "Horas": [horas_totais]})
fig_horas = px.bar(
    df_total_horas,
    x="Indicador",
    y="Horas",
    text="Horas",
    title="TEMPO TOTAL EM HORAS (GERAL)",
    color_discrete_sequence=["#2563EB"]
)
fig_horas.update_traces(textposition="outside")
fig_horas.update_layout(paper_bgcolor="white", plot_bgcolor="white", font_color="#111827", height=460, yaxis_title="Horas", xaxis_title="")

with g2:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.plotly_chart(fig_horas, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)

g3, g4 = st.columns(2)

df_horas = df[[col_titulo, col_horas]].sort_values(col_horas, ascending=True).tail(15)
fig_horas_tarefa = px.bar(
    df_horas,
    x=col_horas,
    y=col_titulo,
    orientation="h",
    text=col_horas,
    title="TEMPO EM HORAS POR TAREFA",
    color_discrete_sequence=["#2563EB"]
)
fig_horas_tarefa.update_layout(paper_bgcolor="white", plot_bgcolor="white", font_color="#111827", height=580, xaxis_title="Horas", yaxis_title="")

with g3:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.plotly_chart(fig_horas_tarefa, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

df_perc = df[[col_titulo, col_percent]].sort_values(col_percent, ascending=True).tail(15)
fig_perc = px.bar(
    df_perc,
    x=col_percent,
    y=col_titulo,
    orientation="h",
    text=col_percent,
    title="% DE CONCLUSÃO POR TAREFA",
    color_discrete_sequence=["#10B981"]
)
fig_perc.update_layout(paper_bgcolor="white", plot_bgcolor="white", font_color="#111827", height=580, xaxis_title="% de conclusão", yaxis_title="")

with g4:
    st.markdown('<div class="chart-box">', unsafe_allow_html=True)
    st.plotly_chart(fig_perc, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

st.markdown("---")
st.subheader("📋 Planilha carregada")
st.dataframe(df, use_container_width=True, height=450)
