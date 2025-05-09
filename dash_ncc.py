# CONEXÃO COM DATABRICKS
from databricks import sql
import os
import pandas as pd
import io
import unicodedata
import pyodbc
import win32
import win32com.client
import matplotlib.pyplot as plt
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go

# Carregando os dados atualizados
df = pd.read_excel("C:/Users/matheus.Archanjo/Desktop/NCC/BASE_NCC_GERAL_LEONARDO/base_ncc_geral_atualizado.xlsx")

# Convertendo datas para datetime
df['data_processado'] = pd.to_datetime(df['data_processado'], errors='coerce')

# Filtros no sidebar
st.sidebar.title("Ciclos")
data_min = df['data_processado'].min()
data_max = df['data_processado'].max()
data_range = st.sidebar.date_input("Período de processamento", [data_min, data_max])

# Título principal
st.title("Dashboard de Movimentação - NCC")

# KPIs - contadores
col1, col2, col3 = st.columns(3)
col1.metric("Migrações", f"{df['flag_migrou'].sum():,}".replace(",", "."))
col2.metric("Cancelamentos", f"{df['flag_cancelou'].sum():,}".replace(",", "."))
col3.metric("Down Ticket", f"{df['flag_down_ticket'].sum():,}".replace(",", "."))

col4, col5, col6 = st.columns(3)
col4.metric("Downgrade", f"{df['flag_downgrade'].sum():,}".replace(",", "."))
col5.metric("Upgrade", f"{df['flag_upgrade'].sum():,}".replace(",", "."))
col6.metric("Suspensos", f"{df['flag_suspenso'].sum():,}".replace(",", "."))

# Gráfico de barras com todas as flags
flag_counts = {
    "Migrou": df['flag_migrou'].sum(),
    "Cancelou": df['flag_cancelou'].sum(),
    "Downgrade": df['flag_downgrade'].sum(),
    "Upgrade": df['flag_upgrade'].sum(),
    "Suspenso": df['flag_suspenso'].sum(),
    "Down Ticket": df['flag_down_ticket'].sum()
}

fig_bar = px.bar(
    x=list(flag_counts.keys()), 
    y=list(flag_counts.values()), 
    labels={'x': 'Flag', 'y': 'Quantidade'}, 
    title="Quantidade por Flag",
    text=[f"{v:,}".replace(",", ".") for v in flag_counts.values()]
)
st.plotly_chart(fig_bar, use_container_width=True)

# Tabela com registros detalhados
st.subheader("Visualizar Dados Detalhados")
st.dataframe(df.head(100))

# Total de registros no período
total = df['codigo_contrato_air'].count()
processados = df['flag_processado'].sum()
cancelados = df['flag_cancelou'].sum()
migrou = df['flag_migrou'].sum()
downgrade = df['flag_downgrade'].sum()
suspenso = df['flag_suspenso'].sum()
downticket = df['flag_down_ticket'].sum()

saldo = processados - cancelados - migrou - downgrade - suspenso - downticket

fig_cascata = go.Figure(go.Waterfall(
    name="Movimentações",
    orientation="v",
    measure=[
        "absolute", "absolute", "relative", "relative", "relative", "relative", "relative", "total"
    ],
    x=[
        "Total", "Processados", "-Cancelados", "-Migrou", "-Downgrade", "-Suspenso", "-Down Ticket", "Saldo"
    ],
    y=[
        total, processados, -cancelados, -migrou, -downgrade, -suspenso, -downticket, saldo
    ],
    connector={"line": {"color": "gray"}},
    decreasing={"marker": {"color": "#EF553B"}},
    totals={"marker": {"color": "#636EFA"}},
))

fig_cascata.update_layout(
    title="Cascata de Movimentações",
    waterfallgap=0.3,
    yaxis_title="Quantidade"
)

st.plotly_chart(fig_cascata, use_container_width=True)

# Tabela da cascata formatada
data = {
    'Indicador': ['Total', 'Processados', 'Cancelados', 'Migrou', 'Downgrade', 'Suspenso', 'Down Ticket'],
    'Valor': [total, processados, cancelados, migrou, downgrade, suspenso, downticket]
}
df_cascata = pd.DataFrame(data)
df_cascata['Valor'] = df_cascata['Valor'].apply(lambda x: f"{x:,.0f}".replace(",", "X").replace(".", ",").replace("X", "."))
st.dataframe(df_cascata)

# Receita por ciclo
receita_nao = df[df['não processado'] == 'false'].groupby('ciclo')['valor_total_destino'].sum()
receita_baixa = df[df['flag_cancelou'] == 1].groupby('ciclo')['ticket_final'].sum()
receita_suspensos = df[df['flag_suspenso'] == 1].groupby('ciclo')['ticket_final'].sum()
receita_comunicados = df[df['flag_comunicado'] == 1].groupby('ciclo')['valor_total_destino'].sum()
receita_total = df[df['flag_processado'] == 1].groupby('ciclo')['valor_total_destino'].sum()

# Combina em um DataFrame
df_final = pd.DataFrame({
    'Receita não processados': receita_nao,
    'Receita Baixa': receita_baixa,
    'Receita Suspensos': receita_suspensos,
    'Receita comunicados': receita_comunicados,
    'Receita total': receita_total
}).fillna(0)

# Formata valores no padrão brasileiro
df_final_formatado = df_final.applymap(lambda x: f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

# Mostra no Streamlit
df_final_formatado.index.name = 'Ciclo'
st.subheader("Receita por Ciclo (R$)")
st.dataframe(df_final_formatado)


# total = df.groupby('ciclo')['codigo_contrato_air'].count()
# comunicados = df.groupby('ciclo')['flag_comunicado'].sum().rename("Qtd Contratos Comunicados")
# nao_processados = df[df['não processado'] == "false"].groupby('ciclo')['não processado'].count().rename("Não processados")
# processados = df.groupby('ciclo')['flag_processado'].sum().rename("Qtd Contratos Processados")
# cancelados = df.groupby('ciclo')['flag_cancelou'].sum().rename("Baixas")
# downgrade = df.groupby('ciclo')['flag_downgrade'].sum().rename("Downgrade")
# upgrade = df.groupby('ciclo')['flag_upgrade'].sum().rename("Upgrade")
# suspensos = df.groupby('ciclo')['flag_suspenso'].sum().rename("Suspensos")
# downticket = df.groupby('ciclo')['flag_down_ticket'].sum().rename("Downgrade Ticket")