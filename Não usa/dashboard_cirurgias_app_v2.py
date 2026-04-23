
import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard Cirurgias CCP - IAVC", layout="wide")
st.title("📊 Dashboard de Cirurgias - CCP (Janeiro à abril de 2025)")

@st.cache_data
def carregar_dados():
    df = pd.read_excel("cirurgias_cp_1ºtrim.xlsx")
    df['DATA'] = pd.to_datetime(df['DATA'], errors='coerce')
    df['ANO_MES'] = df['DATA'].dt.to_period('M').astype(str)
    df['CHEFE'] = df['CHEFE'].str.upper().str.strip()
    df['CIRURGIA_GRUPO'] = df['CIRURGIA_GRUPO'].str.upper().str.strip()
    df['GRUPO_MESTRE'] = df['GRUPO_MESTRE'].str.upper().str.strip()
    return df

df = carregar_dados()

# Layout com colunas
col1, col2 = st.columns(2)

# Gráfico de tendência (linha) para número de procedimentos por mês
with col1:
    st.subheader("📈 Número de Procedimentos por Mês")
    df_mes = df.groupby("ANO_MES").size().reset_index(name="Quantidade")
    fig_linha = px.line(df_mes, x="ANO_MES", y="Quantidade", markers=True)
    fig_linha.update_traces(line_color='royalblue')
    st.plotly_chart(fig_linha, use_container_width=True)

# Cirurgias por grupo com múltiplas cores
with col2:
    st.subheader("🏥 Cirurgias")
    df_grupo = df["CIRURGIA_GRUPO"].value_counts().reset_index()
    df_grupo.columns = ["Tipo de Cirurgia", "Quantidade"]
    fig_grupo = px.bar(df_grupo, x="Quantidade", y="Tipo de Cirurgia", orientation="h", text="Quantidade", color="Tipo de Cirurgia")
    st.plotly_chart(fig_grupo, use_container_width=True)

# Cirurgias por patologia (grupo mestre)
st.subheader("🧠 Cirurgias por Patologia")
df_mestre = df["GRUPO_MESTRE"].value_counts().reset_index()
df_mestre.columns = ["Grupo Patológico", "Quantidade"]
fig_mestre = px.pie(df_mestre, names="Grupo Patológico", values="Quantidade", hole=0.3)
st.plotly_chart(fig_mestre, use_container_width=True)

# Tabela de duração por tipo de cirurgia com estatísticas
st.subheader("⏱️ Estatísticas de Duração por Tipo de Cirurgia")
df_estatisticas = df.groupby("CIRURGIA_GRUPO")["DURACAO_HORAS"].agg(["count", "min", "max", "mean", "std"]).reset_index()
df_estatisticas.columns = ["Tipo de Cirurgia", "N", "Mínimo (h)", "Máximo (h)", "Média (h)", "Desvio Padrão (h)"]
st.dataframe(df_estatisticas)

# Número de cirurgias por chefe com múltiplas cores
st.subheader("👨‍⚕️ Cirurgias por Cirurgião Chefe")
df_chefe = df["CHEFE"].value_counts().reset_index()
df_chefe.columns = ["Chefe", "Quantidade"]
fig_chefe = px.bar(df_chefe, x="Quantidade", y="Chefe", orientation="h", text="Quantidade", color="Chefe")
st.plotly_chart(fig_chefe, use_container_width=True)

# 🔎 Seção de busca por número MV
st.subheader("🔎 Buscar Paciente por Número MV")

# Campo de entrada do usuário
mv_input = st.text_input("Digite o número MV do paciente (exato):")

if mv_input:
    resultado = df[df['MV'].astype(str) == mv_input.strip()]
    if not resultado.empty:
        st.success(f"Encontrado {len(resultado)} registro(s) com MV = {mv_input}")
        st.dataframe(resultado, use_container_width=True, height=600)

        # Adicionar opção de download
        st.download_button(
            label="📥 Baixar resultados em Excel",
            data=resultado.to_excel(index=False, engine='openpyxl'),
            file_name=f"pacientes_mv_{mv_input}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    else:
        st.warning("Nenhum paciente encontrado com esse número MV.")


st.markdown("---")
st.caption("Dashboard desenvolvido por CECI - Computational Excellence for Clinical Innovation")