import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

st.set_page_config(layout="wide")

# Carregar dados (assume-se que o arquivo esteja no mesmo diretório)
df = pd.read_excel("Todas Cirurgias-CCP.xlsx", sheet_name="Plan1")
df.columns = df.columns.str.strip()
df['DATA'] = pd.to_datetime(df['DATA'])
df['CIRURGIA'] = df['CIRURGIA'].str.upper()
df['Duração em horas'] = pd.to_numeric(df['Duração em horas'], errors='coerce')


df.shape

# Agrupamento padronizado
def agrupar_procedimentos(cirurgia):
    import re
    if 'TIREOIDECTOMIA TOTAL' in cirurgia and (
        'LINFADENECTOMIA' in cirurgia or 'ESVAZIAMENTO' in cirurgia or
        'RECORRENCIA' in cirurgia or 'RECORRENCIAL' in cirurgia
    ):
        return 'TIREOIDECTOMIA TOTAL + ESV LN'
    elif 'TIREOIDECTOMIA TOTAL' in cirurgia:
        return 'TIREOIDECTOMIA TOTAL'
    elif 'TIREOIDECTOMIA PARCIAL' in cirurgia:
        return 'TIREOIDECTOMIA PARCIAL'
    elif re.search(r'TROCA.*PR[ÓO]TESE.*FONAT', cirurgia):
        return 'TROCA DE PRÓTESE FONATÓRIA'
    elif re.search(r'RECONSTRU[ÇC][AÃ]O.*FONAT', cirurgia):
        return 'RECONSTRUÇÃO PARA FONAÇÃO'
    return cirurgia

df['Procedimento Agrupado'] = df['CIRURGIA'].apply(agrupar_procedimentos)
df['Ano_Mes'] = df['DATA'].dt.to_period('M').astype(str)

# Sidebar
st.sidebar.header("Filtros")
procedimentos = df['Procedimento Agrupado'].unique()
proc_sel = st.sidebar.multiselect("Procedimentos", options=procedimentos, default=list(procedimentos))

# Filtro de dados
df_filtrado = df[df['Procedimento Agrupado'].isin(proc_sel)]

# Gráfico de evolução mensal
evolucao = df_filtrado.groupby(['Ano_Mes', 'Procedimento Agrupado']).size().reset_index(name='Total')
st.subheader("📈 Evolução Mensal por Procedimento")
fig1, ax1 = plt.subplots(figsize=(10, 5))
sns.lineplot(data=evolucao, x='Ano_Mes', y='Total', hue='Procedimento Agrupado', marker='o', ax=ax1)
plt.xticks(rotation=45)
plt.tight_layout()
st.pyplot(fig1)

# Top 10 procedimentos
top_proc = df_filtrado['Procedimento Agrupado'].value_counts().head(10).reset_index()
top_proc.columns = ['Procedimento', 'Total']
st.subheader("🔝 Top 10 Procedimentos Realizados")
fig2, ax2 = plt.subplots(figsize=(10, 6))
sns.barplot(data=top_proc, y='Procedimento', x='Total', ax=ax2)
st.pyplot(fig2)

# Tabela resumo com tempo médio
df_resumo = df_filtrado.groupby('Procedimento Agrupado').agg(
    Total=('Procedimento Agrupado', 'count'),
    Duração_Média_h=('Duração em horas', 'mean')
).sort_values(by='Total', ascending=False)
st.subheader("📊 Resumo por Procedimento")
st.dataframe(df_resumo, use_container_width=True)

st.markdown("---")
st.caption("Dashboard desenvolvido para análise de cirurgias de cabeça e pescoço - CECI HealthTech")
