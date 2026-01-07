import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(
    page_title="Sistema de Gestão",
    page_icon="📊",
    layout="wide"
)

# --- FUNÇÕES ÚTEIS ---
def init_session_state():
    """Inicia a 'memória' do sistema para guardar dados temporariamente"""
    if 'vendas' not in st.session_state:
        # Dados iniciais de exemplo
        st.session_state['vendas'] = [
            {"Data": datetime(2023, 10, 1), "Vendedor": "João", "Produto": "Notebook", "Valor": 3500.00, "Qtd": 1},
            {"Data": datetime(2023, 10, 2), "Vendedor": "Maria", "Produto": "Mouse", "Valor": 150.00, "Qtd": 5},
            {"Data": datetime(2023, 10, 2), "Vendedor": "João", "Produto": "Teclado", "Valor": 200.00, "Qtd": 2},
        ]

def get_data():
    """Transforma a lista da sessão em um DataFrame pandas"""
    return pd.DataFrame(st.session_state['vendas'])

# --- INICIALIZAÇÃO ---
init_session_state()

# --- SIDEBAR (MENU) ---
st.sidebar.title("Navegação")
page = st.sidebar.radio("Ir para:", ["Dashboard 📈", "Novo Registro 📝", "Base de Dados 📂"])
st.sidebar.markdown("---")
st.sidebar.info("Este sistema roda inteiramente no navegador usando Streamlit.")

# --- PÁGINA 1: DASHBOARD ---
if page == "Dashboard 📈":
    st.title("📊 Dashboard de Vendas")
    
    df = get_data()
    
    if not df.empty:
        # Métricas (KPIs)
        col1, col2, col3 = st.columns(3)
        total_vendas = df["Valor"].sum()
        qtd_total = df["Qtd"].sum()
        ticket_medio = total_vendas / len(df) if len(df) > 0 else 0
        
        col1.metric("Faturamento Total", f"R$ {total_vendas:,.2f}")
        col2.metric("Itens Vendidos", f"{qtd_total}")
        col3.metric("Ticket Médio", f"R$ {ticket_medio:,.2f}")
        
        st.markdown("---")
        
        # Gráficos
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            st.subheader("Vendas por Vendedor")
            fig_vendedor = px.bar(df, x="Vendedor", y="Valor", color="Vendedor", template="plotly_white")
            st.plotly_chart(fig_vendedor, use_container_width=True)
            
        with col_g2:
            st.subheader("Faturamento por Produto")
            fig_produto = px.pie(df, values="Valor", names="Produto", hole=0.4)
            st.plotly_chart(fig_produto, use_container_width=True)
            
    else:
        st.warning("Nenhum dado registrado ainda. Vá para a aba 'Novo Registro'.")

# --- PÁGINA 2: CADASTRO ---
elif page == "Novo Registro 📝":
    st.title("📝 Registrar Nova Venda")
    
    with st.form(key="form_venda", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            data = st.date_input("Data da Venda", datetime.now())
            vendedor = st.selectbox("Vendedor", ["João", "Maria", "Carlos", "Ana"])
            
        with col2:
            produto = st.text_input("Nome do Produto")
            qtd = st.number_input("Quantidade", min_value=1, value=1, step=1)
            valor_unit = st.number_input("Valor Total da Venda (R$)", min_value=0.0, format="%.2f")
            
        submit_button = st.form_submit_button(label="Salvar Venda")
        
        if submit_button:
            if produto and valor_unit > 0:
                nova_venda = {
                    "Data": pd.to_datetime(data),
                    "Vendedor": vendedor,
                    "Produto": produto,
                    "Valor": valor_unit,
                    "Qtd": qtd
                }
                st.session_state['vendas'].append(nova_venda)
                st.success(f"Venda de '{produto}' registrada com sucesso!")
            else:
                st.error("Por favor, preencha o nome do produto e o valor.")

# --- PÁGINA 3: DADOS ---
elif page == "Base de Dados 📂":
    st.title("📂 Histórico de Transações")
    
    df = get_data()
    
    # Filtros simples
    vendedor_filtro = st.multiselect("Filtrar por Vendedor", df["Vendedor"].unique())
    if vendedor_filtro:
        df = df[df["Vendedor"].isin(vendedor_filtro)]
        
    st.dataframe(df, use_container_width=True)
    
    # Botão de Download
    csv = df.to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Baixar CSV",
        data=csv,
        file_name="relatorio_vendas.csv",
        mime="text/csv",
    )
