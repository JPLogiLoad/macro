import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

# --- CONFIGURAÇÃO ---
st.set_page_config(page_title="LogiLoad - Torre de Controle", page_icon="🚚", layout="wide")

# --- DADOS FICTÍCIOS (Simulando seu Supabase) ---
def get_logistics_data():
    # Simulando dados de 50 entregas espalhadas por SP/Brasil
    data = {
        'id_entrega': range(1001, 1051),
        'status': np.random.choice(['Em Trânsito', 'Entregue', 'Atrasado', 'RNC Abertos'], 50, p=[0.4, 0.4, 0.1, 0.1]),
        'lat': np.random.uniform(-23.6, -23.4, 50), # Latitudes próximas a SP
        'lon': np.random.uniform(-46.7, -46.5, 50), # Longitudes próximas a SP
        'motorista': np.random.choice(['Carlos', 'Ana', 'Beto', 'Dina'], 50),
        'valor_frete': np.random.uniform(500, 2500, 50)
    }
    return pd.DataFrame(data)

df = get_logistics_data()

# --- SIDEBAR ---
st.sidebar.image("https://cdn-icons-png.flaticon.com/512/760/760706.png", width=50) # Ícone genérico
st.sidebar.title("LogiLoad Control")
st.sidebar.markdown("Filtros Operacionais")

filtro_status = st.sidebar.multiselect(
    "Filtrar por Status", 
    options=df['status'].unique(),
    default=df['status'].unique()
)

# Aplicar filtro
df_filtrado = df[df['status'].isin(filtro_status)]

# --- DASHBOARD ---
st.title("🚚 LogiLoad - Torre de Controle Operacional")

# 1. KPIs (Indicadores Chave)
col1, col2, col3, col4 = st.columns(4)
total_entregas = len(df_filtrado)
entregas_rnc = len(df_filtrado[df_filtrado['status'] == 'RNC Abertos'])
receita_total = df_filtrado['valor_frete'].sum()

col1.metric("Entregas no Radar", total_entregas)
col2.metric("Ocorrências (RNC)", entregas_rnc, delta=-entregas_rnc, delta_color="inverse")
col3.metric("Veículos em Trânsito", len(df_filtrado[df_filtrado['status'] == 'Em Trânsito']))
col4.metric("Receita de Frete (Visível)", f"R$ {receita_total:,.2f}")

st.markdown("---")

# 2. Mapa e Tabela
c_mapa, c_tabela = st.columns([2, 1])

with c_mapa:
    st.subheader("📍 Rastreamento em Tempo Real")
    # Streamlit tem mapa nativo simples, mas o pydeck ou plotly são mais avançados.
    # Usando st.map para simplicidade:
    st.map(df_filtrado, latitude='lat', longitude='lon', size=20, color='#0044ff')
    st.caption("*Dados de localização simulados para demonstração.")

with c_tabela:
    st.subheader("📋 Lista de Cargas")
    # Destacar as linhas com problema (RNC ou Atraso)
    st.dataframe(
        df_filtrado[['id_entrega', 'status', 'motorista', 'valor_frete']],
        hide_index=True,
        use_container_width=True,
        height=400
    )

st.markdown("---")

# 3. Análise de Performance
st.subheader("📊 Análise de Performance")
col_g1, col_g2 = st.columns(2)

with col_g1:
    fig_status = px.pie(df, names='status', title='Distribuição de Status das Entregas', hole=0.4)
    st.plotly_chart(fig_status, use_container_width=True)

with col_g2:
    # Agrupando faturamento por motorista
    df_fat = df.groupby('motorista')['valor_frete'].sum().reset_index()
    fig_bar = px.bar(df_fat, x='motorista', y='valor_frete', title='Faturamento por Motorista', color='valor_frete')
    st.plotly_chart(fig_bar, use_container_width=True)
