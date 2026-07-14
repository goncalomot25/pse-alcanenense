
import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import plotly.express as px
import base64

st.set_page_config(
    page_title="PSE Alcanenense",
    page_icon="⚽",
    layout="wide"
)

DATA_FILE = Path("respostas_pse.csv")
LOGO_FILE = Path("logo_aca.png")
SCALE_FILE = Path("escala_pse.png")
PASSWORD_TREINADOR = "aca2026"

# Lista de atletas para aparecer na drop-list.
# Podes alterar estes nomes no GitHub quando tiveres a lista final.
ATLETAS = [
    "Seleciona o teu nome",
    "Afonso Maia Netto",
    "Afonso Quaresma Bento",
    "Andre Afonso Boura",
    "Andre Filipe Martins Fernandes",
    "Diogo Ferreira Dos Santos Paveia",
    "Duarte Ribeiro Sousa Henriques",
    "Gonçalo Filipe Rosa Mota",
    "Henrique Pereira Damásio",
    "Hugo Pereira Carvalho",
    "Lucas Nunes Freire",
    "Martim Ferreira Moreira",
    "Miguel Vila Nova Rocha",
    "Nuno Gabriel Souto Silva",
    "Pedro De Sousa Mota Gameiro",
    "Ricardo Miguel Sousa Faria",
    "Ryan Mateus Villa Nova Silveira",
    "Simão Fernandes Sousa",
    "Tiago Roberto Rego Teixeira",
    "Outro"
]


def carregar_dados():
    if DATA_FILE.exists():
        return pd.read_csv(DATA_FILE)
    return pd.DataFrame(columns=[
        "Data e hora", "Data", "Hora", "Primeiro nome",
        "Segundo nome", "Atleta", "PSE", "Classificação"
    ])


def guardar_resposta(resposta):
    df = carregar_dados()
    df = pd.concat([df, pd.DataFrame([resposta])], ignore_index=True)
    df.to_csv(DATA_FILE, index=False)


def classificar_pse(valor):
    if valor == 0:
        return "Repouso"
    if valor == 1:
        return "Muito, muito fácil"
    if valor == 2:
        return "Fácil"
    if valor == 3:
        return "Moderado"
    if valor == 4:
        return "Algo difícil"
    if valor in [5, 6]:
        return "Difícil"
    if valor in [7, 8, 9]:
        return "Muito difícil"
    return "Esforço máximo"


def nivel_alerta(valor):
    if valor <= 3:
        return "baixo"
    if valor <= 6:
        return "moderado"
    return "elevado"


def preparar_excel(df):
    colunas_excel = ["Data e hora", "Atleta", "PSE", "Classificação"]
    df_excel = df[colunas_excel].copy()
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_excel.to_excel(writer, index=False, sheet_name="Respostas PSE")
    output.seek(0)
    return output


st.markdown("""
<style>
    .stApp {
        background: #f5e51b;
    }

    [data-testid="stSidebar"] {
        background-color: #111111;
    }

    [data-testid="stSidebar"] * {
        color: #f5e51b !important;
    }

    .block-container {
        max-width: 1180px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    .header-box {
        background-color: #050505;
        border-radius: 22px;
        padding: 28px 20px 32px 20px;
        text-align: center;
        margin-bottom: 28px;
        box-shadow: 0px 8px 24px rgba(0,0,0,0.25);
    }

    .header-title {
        color: #f5e51b;
        font-size: 34px;
        font-weight: 900;
        margin-top: 10px;
        margin-bottom: 6px;
        line-height: 1.15;
    }

    .header-subtitle {
        color: white;
        font-size: 16px;
        font-weight: 500;
    }

    h1, h2, h3 {
        color: #111111;
        font-weight: 850;
    }

    .stApp, .stApp p, .stApp label, .stApp span, .stApp div {
        color: #111111;
    }

    .header-box, .header-box div, .header-title {
        color: #f5e51b !important;
    }

    .header-subtitle {
        color: #ffffff !important;
    }

    [data-testid="stSidebar"], [data-testid="stSidebar"] * {
        color: #f5e51b !important;
    }

    .stAlert div {
        color: #111111 !important;
    }

    .stButton > button, .stDownloadButton > button {
        background-color: #050505;
        color: #f5e51b !important;
        border: none;
        border-radius: 12px;
        font-weight: 800;
        padding: 0.65rem 1.4rem;
    }

    .stButton > button:hover, .stDownloadButton > button:hover {
        background-color: #222222;
        color: #fff27a !important;
    }

    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border-radius: 16px;
        padding: 14px;
        box-shadow: 0px 5px 16px rgba(0,0,0,0.10);
        border-left: 6px solid #111111;
    }
</style>
""", unsafe_allow_html=True)

logo_html = ""
if LOGO_FILE.exists():
    logo_bytes = LOGO_FILE.read_bytes()
    logo_b64 = base64.b64encode(logo_bytes).decode()
    logo_html = f'<img src="data:image/png;base64,{logo_b64}" style="width:105px; border-radius:12px; margin-bottom:8px;">'

st.markdown(
    f"""
    <div class="header-box">
        {logo_html}
        <div class="header-title">Formulário de Perceção Subjetiva de Esforço</div>
        <div class="header-subtitle">Atlético Clube Alcanenense</div>
    </div>
    """,
    unsafe_allow_html=True
)

pagina = st.sidebar.radio(
    "Escolhe a zona:",
    ["Atleta - Responder", "Treinador - Análise"]
)

if pagina == "Atleta - Responder":
    st.subheader("Escala PSE")

    if SCALE_FILE.exists():
        st.image(str(SCALE_FILE), use_container_width=True)
    else:
        st.warning("A imagem da escala PSE não foi encontrada.")

    st.divider()

    st.subheader("Responder ao formulário")
    st.write("Seleciona o teu nome e indica o número que melhor representa o esforço que sentiste durante o treino.")

    with st.form("formulario_pse", clear_on_submit=True):
        atleta_selecionado = st.selectbox("Nome do atleta", ATLETAS)

        atleta_outro = ""
        if atleta_selecionado == "Outro":
            atleta_outro = st.text_input("Escreve o teu nome completo")

        pse = st.select_slider(
            "Qual foi a tua perceção do esforço durante o treino de hoje?",
            options=list(range(0, 11)),
            value=5
        )

        st.info(f"Valor selecionado: {pse} — {classificar_pse(pse)}")

        enviar = st.form_submit_button("Enviar resposta")

        if enviar:
            if atleta_selecionado == "Seleciona o teu nome":
                st.error("Seleciona o teu nome.")
            elif atleta_selecionado == "Outro" and atleta_outro.strip() == "":
                st.error("Escreve o teu nome completo.")
            else:
                agora = datetime.now()
                atleta = atleta_outro.strip() if atleta_selecionado == "Outro" else atleta_selecionado

                partes_nome = atleta.split()
                primeiro_nome = partes_nome[0] if len(partes_nome) > 0 else ""
                segundo_nome = " ".join(partes_nome[1:]) if len(partes_nome) > 1 else ""

                resposta = {
                    "Data e hora": agora.strftime("%Y-%m-%d %H:%M:%S"),
                    "Data": agora.strftime("%Y-%m-%d"),
                    "Hora": agora.strftime("%H:%M:%S"),
                    "Primeiro nome": primeiro_nome,
                    "Segundo nome": segundo_nome,
                    "Atleta": atleta,
                    "PSE": pse,
                    "Classificação": classificar_pse(pse)
                }

                guardar_resposta(resposta)
                st.success("Resposta enviada com sucesso. Obrigado!")

    st.divider()

    st.subheader("Interpretação rápida")
    col_i1, col_i2, col_i3 = st.columns(3)

    with col_i1:
        st.write("**0** — Repouso")
        st.write("**1-2** — Muito fácil / Fácil")

    with col_i2:
        st.write("**3-4** — Moderado / Algo difícil")
        st.write("**5-6** — Difícil")

    with col_i3:
        st.write("**7-9** — Muito difícil")
        st.write("**10** — Esforço máximo")


if pagina == "Treinador - Análise":
    with st.container(border=True):
        st.subheader("Área dos treinadores")
        st.write("Acesso reservado aos treinadores para consulta, análise e exportação dos dados.")

        password = st.text_input("Palavra-passe dos treinadores", type="password")

        if password != PASSWORD_TREINADOR:
            st.warning("Insere a palavra-passe para aceder à análise.")
            st.stop()

        df = carregar_dados()

        if df.empty:
            st.info("Ainda não existem respostas.")
            st.stop()

        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.date

        st.success("Acesso autorizado.")

        st.markdown("### Filtros")
        f1, f2 = st.columns(2)

        datas = sorted(df["Data"].dropna().unique())
        with f1:
            data_escolhida = st.selectbox(
                "Filtrar por data",
                ["Todas"] + [str(d) for d in datas]
            )

        atletas = sorted(df["Atleta"].dropna().unique())
        with f2:
            atleta_escolhido = st.selectbox(
                "Filtrar por atleta",
                ["Todos"] + atletas
            )

        df_f = df.copy()

        if data_escolhida != "Todas":
            df_f = df_f[df_f["Data"].astype(str) == data_escolhida]

        if atleta_escolhido != "Todos":
            df_f = df_f[df_f["Atleta"] == atleta_escolhido]

        if df_f.empty:
            st.warning("Não existem dados para estes filtros.")
            st.stop()

        st.markdown("### Resumo automático")

        media_pse = df_f["PSE"].mean()
        max_pse = df_f["PSE"].max()
        min_pse = df_f["PSE"].min()
        total = len(df_f)
        elevados = df_f[df_f["PSE"] >= 7]

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Respostas", total)
        m2.metric("Média PSE", f"{media_pse:.1f}")
        m3.metric("PSE mais alta", int(max_pse))
        m4.metric("PSE mais baixa", int(min_pse))

        st.write(
            f"Foram registadas {total} respostas. A média da PSE foi {media_pse:.1f}, "
            f"o que representa um nível {nivel_alerta(media_pse)} de esforço percebido. "
            f"A PSE máxima foi {int(max_pse)} e a mínima foi {int(min_pse)}."
        )

        if not elevados.empty:
            st.warning(
                "Atletas com PSE elevada: "
                + ", ".join(elevados["Atleta"].tolist())
            )
        else:
            st.info("Não existem atletas com PSE elevada neste filtro.")

        tab1, tab2, tab3 = st.tabs(["Tabela", "Gráficos", "Exportar Excel"])

        with tab1:
            st.dataframe(df_f[["Data e hora", "Atleta", "PSE", "Classificação"]], use_container_width=True, hide_index=True)

        with tab2:
            fig = px.bar(
                df_f,
                x="Atleta",
                y="PSE",
                text="PSE",
                title="PSE por atleta",
                color="PSE",
                color_continuous_scale=[
                    (0.0, "#f2f2f2"),
                    (0.2, "#f0e48c"),
                    (0.5, "#ffd000"),
                    (0.8, "#ff8a00"),
                    (1.0, "#e40012")
                ]
            )
            fig.update_layout(yaxis_range=[0, 10], coloraxis_showscale=False)
            st.plotly_chart(fig, use_container_width=True)

            if df_f["Data"].nunique() > 1:
                evolucao = df_f.groupby("Data", as_index=False)["PSE"].mean()
                fig2 = px.line(
                    evolucao,
                    x="Data",
                    y="PSE",
                    markers=True,
                    title="Evolução da média da PSE"
                )
                fig2.update_layout(yaxis_range=[0, 10])
                st.plotly_chart(fig2, use_container_width=True)

        with tab3:
            excel = preparar_excel(df_f)
            st.download_button(
                "Descarregar Excel",
                data=excel,
                file_name="respostas_pse.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
