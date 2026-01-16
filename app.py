import pandas as pd
import streamlit as st
import io
import string
from datetime import datetime, timedelta, time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Cargas Pro", page_icon="🚀", layout="wide")

# --- FUNÇÕES ---

def get_col_letter(n):
    """Converte 1->A, 2->B para ajudar o usuário visualmente."""
    try:
        string_val = ""
        n = int(n)
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string_val = chr(65 + remainder) + string_val
        return string_val
    except:
        return "?"

def carregar_dados_blindado(uploaded_file):
    bytes_data = uploaded_file.getvalue()
    
    # 1. Tentar Excel Padrão (.xlsx)
    try:
        return pd.read_excel(io.BytesIO(bytes_data), header=None)
    except:
        pass
    
    # 2. Tentar Excel Antigo (.xls)
    try:
        return pd.read_excel(io.BytesIO(bytes_data), header=None, engine='xlrd')
    except:
        pass

    # 3. Tentar Texto/CSV (UTF-16 é comum em SAP)
    encodings = ['utf-16', 'utf-8', 'latin-1', 'cp1252']
    for encoding in encodings:
        try:
            text = bytes_data.decode(encoding)
            # Separação por TAB (comum SAP)
            try:
                df = pd.read_csv(io.StringIO(text), sep='\t', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass
            # Separação por Ponto e Vírgula
            try:
                df = pd.read_csv(io.StringIO(text), sep=';', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass
        except:
            continue
            
    return None

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações")
    st.info("Agora a contagem é igual ao Excel: A=1, B=2...")
    
    st.subheader("1. Data do Plantão")
    data_ref = st.date_input("Data de Início", datetime.now().date())
    
    st.divider()
    
    st.subheader("2. Mapeamento de Colunas (A=1)")
    # Valores ajustados para usar contagem humana (1, 2, 3...)
    idx_data = st.number_input("DATA (L=12)", value=12, min_value=1, help="Coluna L é a 12ª")
    idx_local = st.number_input("LOCAL (E=5)", value=5, min_value=1, help="Coluna E é a 5ª")
    idx_uf = st.number_input("UF (I=9)", value=9, min_value=1, help="Coluna I é a 9ª")
    idx_transp = st.number_input("TRANSP (K=11)", value=11, min_value=1, help="Coluna K é a 11ª")
    idx_status = st.number_input("STATUS (P=16)", value=16, min_value=1, help="Coluna P é a 16ª")

# --- CORPO PRINCIPAL ---
st.title("🚀 Processador de Cargas Pro")

uploaded_file = st.file_uploader("Arraste seu arquivo aqui", type=["xls", "xlsx", "xlsm", "csv", "txt"])

if uploaded_file:
    df_raw = carregar_dados_blindado(uploaded_file)

    if df_raw is None:
        st.error("❌ Não foi possível ler o arquivo. Formato irreconhecível.")
        st.stop()

    # Ajuste de índice (Humano [1] -> Python [0])
    p_data = idx_data - 1
    p_local = idx_local - 1
    p_uf = idx_uf - 1
    p_transp = idx_transp - 1
    p_status = idx_status - 1

    # --- MAPEADOR VISUAL ---
    with st.expander("🕵️‍♀️ Visualizador de Colunas", expanded=True):
        st.write("Confira se a coluna Data mostra datas reais ou '#######':")
        preview = df_raw.head(3).T.reset_index()
        preview.columns = ["Índice Python", "Linha 1", "Linha 2", "Linha 3"]
        preview.insert(1, "Coluna Excel", [get_col_letter(i+1) for i in preview["Índice Python"]])
        st.dataframe(preview, height=300, use_container_width=True)

    st.divider()

    # --- PROCESSAMENTO ---
    try:
        # Tratamento de Data
        col_amostra = df_raw[p_data].astype(str)
        if col_amostra.str.contains("####").any():
            st.error(f"🚨 PROBLEMA CRÍTICO: A coluna {get_col_letter(idx_data)} contém '#######' no lugar das datas!")
            st.warning("Isso acontece quando o relatório é exportado do SAP com a coluna muito estreita. Você precisa exportar o arquivo novamente garantindo que a data esteja visível.")
            st.stop()

        df_raw[p_data] = pd.to_datetime(df_raw[p_data], dayfirst=True, errors='coerce')
        df_limpo = df_raw.dropna(subset=[p_data]).copy()
        
        min_dt = df_limpo[p_data].min()
        max_dt = df_limpo[p_data].max()
        
        if pd.isna(min_dt):
            st.warning("⚠️ Nenhuma data válida encontrada. Verifique o mapeamento.")
            st.stop()
            
        st.success(f"✅ Arquivo processado! Datas de {min_dt} até {max_dt}")

        # Filtros
        inicio = datetime.combine(data_ref, time(17, 0))
        fim = datetime.combine(data_ref + timedelta(days=1), time(7, 0))
        
        # 1. Data
        df_f1 = df_limpo[(df_limpo[p_data] >= inicio) & (df_limpo[p_data] <= fim)]
        
        # 2. Local
        locais = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
        df_f2 = df_f1[df_f1[p_local].astype(str).str.strip().str.upper().isin(locais)]
        
        # 3. Status
        status_ok = ["SILVER", "GOLD", "DIAMOND"]
        df_f3 = df_f2[df_f2[p_status].astype(str).str.strip().str.upper().isin(status_ok)]
        
        # 4. Regra MG
        mask_mg = ~((df_f3[p_uf].astype(str).str.strip().str.upper() == "MG") & 
                    (df_f3[p_status].astype(str).str.strip().str.upper() == "SILVER"))
        df_f4 = df_f3[mask_mg]
        
        # 5. Transp
        transp_block = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
        df_final = df_f4[~df_f4[p_transp].astype(str).str.strip().str.upper().isin(transp_block)]

        st.write("### 📉 Funil de Resultados")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("1. Data", len(df_f1))
        c2.metric("2. Local", len(df_f2))
        c3.metric("3. Status", len(df_f3))
        c4.metric("4. MG", len(df_f4))
        c5.metric("5. Final", len(df_final))

        if len(df_final) > 0:
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0] # Indices originais para manter compatibilidade visual se possível, ou ajustar lógica de drop
            # Simplificação: Removemos colunas baseadas no índice Python
            # Ajustando para remover as mesmas colunas relativas (VBA remove V, U... A)
            # Vamos apenas exportar o resultado limpo
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='Relatorio')
                wb = writer.book
                ws = writer.sheets['Relatorio']
                fmt_geral = wb.add_format({'border': 1})
                ws.conditional_format(0, 0, len(df_final)-1, len(df_final.columns)-1, {'type': 'no_blanks', 'format': fmt_geral})
            
            st.download_button("📥 Baixar Planilha", output.getvalue(), "Cargas_Filtradas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")

    except Exception as e:
        st.error(f"Erro: {e}")
