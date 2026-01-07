import pandas as pd
import streamlit as st
import io
from datetime import datetime, timedelta, time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Cargas Pro", page_icon="🚀", layout="wide")

# --- FUNÇÕES AUXILIARES (ENGINEERING) ---

def carregar_dados_blindado(uploaded_file):
    """
    Tenta ler o arquivo de todas as formas possíveis (XLS, XLSX, HTML, CSV, TXT).
    Retorna um DataFrame limpo ou None se falhar.
    """
    df = None
    bytes_data = uploaded_file.getvalue()
    
    # Lista de estratégias de leitura
    estrategias = [
        # 1. Excel Padrão
        lambda: pd.read_excel(uploaded_file, header=None),
        # 2. Excel Antigo (xlrd)
        lambda: pd.read_excel(uploaded_file, header=None, engine='xlrd'),
        # 3. HTML (detectando encoding automaticamente)
        lambda: pd.read_html(io.StringIO(bytes_data.decode('utf-8')), header=None)[0],
        lambda: pd.read_html(io.StringIO(bytes_data.decode('latin-1')), header=None)[0],
        lambda: pd.read_html(io.StringIO(bytes_data.decode('cp1252')), header=None)[0],
        # 4. CSV/Texto (Tabulação ou Ponto e Vírgula)
        lambda: pd.read_csv(io.StringIO(bytes_data.decode('latin-1')), sep='\t', header=None, engine='python'),
        lambda: pd.read_csv(io.StringIO(bytes_data.decode('latin-1')), sep=';', header=None, engine='python')
    ]

    for tentativa in estrategias:
        try:
            uploaded_file.seek(0) # Reinicia o arquivo para a próxima tentativa
            df = tentativa()
            if df is not None and not df.empty and df.shape[1] > 5: # Verifica se leu algo útil
                return df
        except Exception:
            continue
    
    return None

def formatar_excel(df):
    """Gera o arquivo Excel final com formatação profissional."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, header=False, sheet_name='Relatorio')
        workbook = writer.book
        worksheet = writer.sheets['Relatorio']
        
        # Formatos
        fmt_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        fmt_geral = workbook.add_format({'border': 1})
        fmt_data = workbook.add_format({'num_format': 'dd/mm/yyyy hh:mm', 'border': 1})
        
        # Aplica bordas em tudo
        if len(df) > 0:
            worksheet.conditional_format(0, 0, len(df)-1, len(df.columns)-1, 
                                        {'type': 'no_blanks', 'format': fmt_geral})
        
        # Ajuste de larguras (estimativa)
        worksheet.set_column(0, 50, 15) 
        
        # Tenta formatar a coluna de valor (agora índice 5 / F) se existir
        if len(df.columns) > 5:
            worksheet.set_column(5, 5, 18, fmt_moeda)

    return output.getvalue()

# --- INTERFACE DO USUÁRIO ---

st.title("🚀 Processador de Cargas Inteligente")
st.markdown("""
Esta ferramenta processa relatórios de cargas aplicando as regras de negócio:
* **Horário:** 17:00 (Dia X) até 07:00 (Dia X+1)
* **Status:** Silver, Gold, Diamond (com exceção de MG+Silver)
* **Local:** CD Pouso Alegre / HPC
""")

# --- BARRA LATERAL (CONFIGURAÇÕES) ---
with st.sidebar:
    st.header("⚙️ Configurações")
    
    st.subheader("1. Data do Plantão")
    # Padrão para hoje, mas permite mudar
    data_ref = st.date_input("Data de Início", datetime.now().date())
    
    st.subheader("2. Mapeamento de Colunas")
    st.info("Ajuste aqui se a planilha mudar de layout.")
    # Índices baseados no VBA original (0 = A, 1 = B...)
    idx_data = st.number_input("Índice Coluna DATA (L)", value=11, min_value=0)
    idx_local = st.number_input("Índice Coluna LOCAL (E)", value=4, min_value=0)
    idx_uf = st.number_input("Índice Coluna UF (I)", value=8, min_value=0)
    idx_transp = st.number_input("Índice Coluna TRANSP. (K)", value=10, min_value=0)
    idx_status = st.number_input("Índice Coluna STATUS (P)", value=15, min_value=0)

# --- CORPO PRINCIPAL ---

uploaded_file = st.file_uploader("Arraste seu arquivo aqui (Excel, HTML, CSV)", type=["xls", "xlsx", "xlsm", "csv", "txt"])

if uploaded_file:
    with st.spinner("Lendo e analisando arquivo..."):
        df_raw = carregar_dados_blindado(uploaded_file)

    if df_raw is None:
        st.error("❌ Não foi possível ler o arquivo. O formato é irreconhecível.")
    else:
        # --- ANÁLISE INICIAL ---
        st.write("---")
        col_metric1, col_metric2, col_metric3 = st.columns(3)
        col_metric1.metric("Linhas Importadas", len(df_raw))
        
        # Tratamento de Data (O mais crítico)
        try:
            # dayfirst=True é essencial para DD/MM/AAAA
            df_raw[idx_data] = pd.to_datetime(df_raw[idx_data], dayfirst=True, errors='coerce')
            
            # Remove linhas onde data é NaT (cabeçalhos repetidos ou lixo)
            df_limpo = df_raw.dropna(subset=[idx_data]).copy()
            
            # Mostra datas encontradas para conferência
            min_dt = df_limpo[idx_data].min()
            max_dt = df_limpo[idx_data].max()
            col_metric2.metric("Menor Data Encontrada", f"{min_dt.day}/{min_dt.month} {min_dt.hour}h")
            col_metric3.metric("Maior Data Encontrada", f"{max_dt.day}/{max_dt.month} {max_dt.hour}h")
            
        except Exception as e:
            st.error(f"Erro ao processar datas na coluna {idx_data}. Verifique o Mapeamento na barra lateral.")
            st.stop()

        # --- APLICAÇÃO DOS FILTROS (PIPELINE) ---
        
        # Definição do Range de Horário
        inicio = datetime.combine(data_ref, time(17, 0))
        fim = datetime.combine(data_ref + timedelta(days=1), time(7, 0))
        
        st.info(f"🔎 Filtrando período: **{inicio}** até **{fim}**")

        # Filtro 1: Data
        mask_data = (df_limpo[idx_data] >= inicio) & (df_limpo[idx_data] <= fim)
        df_f1 = df_limpo[mask_data]
        
        # Filtro 2: Local
        locais = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
        mask_local = df_f1[idx_local].astype(str).str.strip().str.upper().isin(locais)
        df_f2 = df_f1[mask_local]
        
        # Filtro 3: Status
        status_ok = ["SILVER", "GOLD", "DIAMOND"]
        mask_status = df_f2[idx_status].astype(str).str.strip().str.upper().isin(status_ok)
        df_f3 = df_f2[mask_status]
        
        # Filtro 4: Regra MG + Silver (Remove se for MG E Silver)
        # Atenção: A lógica do usuário era "Excluir Silver de MG".
        # Então mantemos se NÃO FOR (MG E SILVER).
        uf_str = df_f3[idx_uf].astype(str).str.strip().str.upper()
        st_str = df_f3[idx_status].astype(str).str.strip().str.upper()
        mask_mg = ~((uf_str == "MG") & (st_str == "SILVER"))
        df_f4 = df_f3[mask_mg]
        
        # Filtro 5: Transportadoras
        transp_block = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
        mask_transp = ~df_f4[idx_transp].astype(str).str.strip().str.upper().isin(transp_block)
        df_final_raw = df_f4[mask_transp]

        # --- EXIBIÇÃO DO FUNIL DE DADOS ---
        st.write("### 📉 Funil de Processamento")
        col_f1, col_f2, col_f3, col_f4, col_f5 = st.columns(5)
        col_f1.metric("Após Data", len(df_f1), delta=len(df_f1)-len(df_limpo))
        col_f2.metric("Após Local", len(df_f2), delta=len(df_f2)-len(df_f1))
        col_f3.metric("Após Status", len(df_f3), delta=len(df_f3)-len(df_f2))
        col_f4.metric("Regra MG", len(df_f4), delta=len(df_f4)-len(df_f3))
        col_f5.metric("Final", len(df_final_raw), delta=len(df_final_raw)-len(df_f4))

        if len(df_final_raw) == 0:
            st.warning("⚠️ O resultado está vazio. Verifique se a 'Data de Início' na barra lateral corresponde às datas do arquivo.")
        else:
            st.success(f"✅ Processamento concluído! {len(df_final_raw)} linhas prontas para download.")

            # --- PREPARAÇÃO FINAL (REMOÇÃO DE COLUNAS) ---
            # VBA Remove: V(21), U(20)... A(0)
            # Vamos remover pelo indice para garantir
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
            # Filtra apenas as que existem (segurança)
            cols_existentes = [c for c in cols_to_drop if c in df_final_raw.columns]
            df_export = df_final_raw.drop(columns=cols_existentes)
            
            # Botão de Download
            arquivo_excel = formatar_excel(df_export)
            
            st.download_button(
                label="📥 Baixar Planilha Processada",
                data=arquivo_excel,
                file_name=f"Cargas_Filtradas_{data_ref}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
            
            with st.expander("🔍 Ver Prévia dos Dados Finais"):
                st.dataframe(df_export.head(10))
