
import streamlit as st
import pandas as pd
import plotly.express as px

# Título
st.set_page_config(page_title="Dashboard Cirurgias CCP - IAVC", layout="wide")
st.title("📊 Dashboard de Cirurgias - CCP (1º trimestre de 2025)")

# Carregamento dos dados
@st.cache_data
def carregar_dados():
    df = pd.read_excel("cirurgias_cp_1ºtrim.xlsx")
    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')
    df['ANO_MES'] = df['DATA'].dt.to_period('M').astype(str)
    df['CHEFE'] = df['CHEFE'].str.upper().str.strip()
    df['CIRURGIA_GRUPO'] = df['CIRURGIA_GRUPO'].str.upper().str.strip()
    df['GRUPO_MESTRE'] = df['GRUPO_MESTRE'].str.upper().str.strip()
    df['ANEST'] = df['ANEST'].str.upper().str.strip()
    return df

df = carregar_dados()

# Layout com colunas
col1, col2 = st.columns(2)

# Número de procedimentos por mês
with col1:
    st.subheader("📅 Número de Procedimentos por Mês")
    df_mes = df.groupby("ANO_MES").size().reset_index(name="Quantidade")
    fig_mes = px.bar(df_mes, x="ANO_MES", y="Quantidade", text="Quantidade")
    st.plotly_chart(fig_mes, use_container_width=True)

# Cirurgias por grupo
with col2:
    st.subheader("🏥 Cirurgias")
    df_grupo = df["CIRURGIA_GRUPO"].value_counts().reset_index()
    df_grupo.columns = ["Tipo de Cirurgia", "Quantidade"]
    fig_grupo = px.bar(df_grupo, x="Quantidade", y="Tipo de Cirurgia", orientation="h", text="Quantidade")
    st.plotly_chart(fig_grupo, use_container_width=True)

# Cirurgias por patologia (grupo mestre)
st.subheader("🧠 Cirurgias por Patologia")
df_mestre = df["GRUPO_MESTRE"].value_counts().reset_index()
df_mestre.columns = ["Grupo Patológico", "Quantidade"]
fig_mestre = px.pie(df_mestre, names="Grupo Patológico", values="Quantidade", hole=0.3)
st.plotly_chart(fig_mestre, use_container_width=True)


# Duração média por tipo de cirurgia
st.subheader("⏱️ Duração Média das Cirurgias (em horas)")
df_duracao = df.groupby("CIRURGIA_GRUPO")["DURACAO_HORAS"].mean().reset_index()
df_duracao.columns = ["Tipo de Cirurgia", "Duração Média (h)"]
fig_duracao = px.bar(df_duracao, x="Duração Média (h)", y="Tipo de Cirurgia", orientation="h", text="Duração Média (h)")
st.plotly_chart(fig_duracao, use_container_width=True)

# Número de cirurgias por chefe
st.subheader("👨‍⚕️ Cirurgias por Cirurgião Chefe")
df_chefe = df["CHEFE"].value_counts().reset_index()
df_chefe.columns = ["Chefe", "Quantidade"]
fig_chefe = px.bar(df_chefe, x="Quantidade", y="Chefe", orientation="h", text="Quantidade")
st.plotly_chart(fig_chefe, use_container_width=True)
