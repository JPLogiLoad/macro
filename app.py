import pandas as pd
import streamlit as st
import io
import string
from datetime import datetime, timedelta, time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Cargas Pro", page_icon="🚀", layout="wide")

# --- FUNÇÕES AUXILIARES ---

def carregar_dados_blindado(uploaded_file):
    """Lê Excel, HTML ou CSV tentando várias estratégias."""
    bytes_data = uploaded_file.getvalue()
    
    # 1. Tentar Excel Padrão (XLSX/XLS)
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, header=None)
    except:
        pass
        
    # 2. Tentar Excel Antigo (xlrd)
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, header=None, engine='xlrd')
    except:
        pass

    # 3. Tentar HTML ou Texto (Detectando encoding)
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            text = bytes_data.decode(encoding)
            # Tenta ler como HTML
            try:
                dfs = pd.read_html(io.StringIO(text), header=None)
                if dfs: return dfs[0]
            except:
                pass
            
            # Tenta ler como CSV/TXT (Tabulação)
            try:
                df = pd.read_csv(io.StringIO(text), sep='\t', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass
                
            # Tenta ler como CSV/TXT (Ponto e Vírgula)
            try:
                df = pd.read_csv(io.StringIO(text), sep=';', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass
        except:
            continue
            
    return None

def get_col_letter(n):
    """Converte índice 0 em letra A, 1 em B..."""
    try:
        return string.ascii_uppercase[n]
    except:
        return "?"

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações")
    
    st.subheader("1. Data do Plantão")
    data_ref = st.date_input("Data de Início", datetime.now().date())
    
    st.divider()
    
    st.subheader("2. Mapeamento de Colunas")
    st.info("Use o 'Visualizador' na tela principal para confirmar os números.")
    idx_data = st.number_input("Índice Coluna DATA (Ex: 11=L)", value=11, min_value=0)
    idx_local = st.number_input("Índice Coluna LOCAL (Ex: 4=E)", value=4, min_value=0)
    idx_uf = st.number_input("Índice Coluna UF (Ex: 8=I)", value=8, min_value=0)
    idx_transp = st.number_input("Índice Coluna TRANSP (Ex: 10=K)", value=10, min_value=0)
    idx_status = st.number_input("Índice Coluna STATUS (Ex: 15=P)", value=15, min_value=0)

# --- CORPO PRINCIPAL ---
st.title("🚀 Processador de Cargas Pro")

uploaded_file = st.file_uploader("Arraste seu arquivo aqui", type=["xls", "xlsx", "xlsm", "csv", "txt"])

if uploaded_file:
    df_raw = carregar_dados_blindado(uploaded_file)

    if df_raw is None:
        st.error("❌ Formato de arquivo desconhecido.")
    else:
        # --- FERRAMENTA VISUAL DE MAPEAMENTO ---
        with st.expander("🕵️‍♀️ Visualizador de Colunas (Use isso para configurar a lateral)", expanded=True):
            st.write("Confira abaixo qual número corresponde à sua coluna:")
            
            # Cria uma tabelinha de prévia transposta para facilitar leitura
            preview = df_raw.head(3).T
            preview.columns = [f"Linha {i+1}" for i in range(3)]
            preview.insert(0, "Índice (Python)", preview.index)
            
            # Adiciona Letra do Excel (A, B, C...)
            preview.insert(1, "Letra Excel", [get_col_letter(i) for i in preview.index])
            
            st.dataframe(preview, use_container_width=True, height=300)

        # --- PROCESSAMENTO ---
        st.divider()
        
        # Tratamento de Data
        try:
            df_raw[idx_data] = pd.to_datetime(df_raw[idx_data], dayfirst=True, errors='coerce')
            df_limpo = df_raw.dropna(subset=[idx_data]).copy()
            
            # Métricas
            min_dt = df_limpo[idx_data].min()
            max_dt = df_limpo[idx_data].max()
            
            col1, col2 = st.columns(2)
            if pd.isna(min_dt):
                col1.error("⚠️ Nenhuma data válida encontrada na coluna selecionada.")
            else:
                col1.success(f"📅 Datas encontradas de: {min_dt.strftime('%d/%m %H:%M')}")
                col2.success(f"Até: {max_dt.strftime('%d/%m %H:%M')}")

        except Exception as e:
            st.error(f"Erro na coluna de data: {e}")
            st.stop()

        # Filtros
        inicio = datetime.combine(data_ref, time(17, 0))
        fim = datetime.combine(data_ref + timedelta(days=1), time(7, 0))
        
        # 1. Data
        mask_data = (df_limpo[idx_data] >= inicio) & (df_limpo[idx_data] <= fim)
        df_f1 = df_limpo[mask_data]
        
        # 2. Local
        locais = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
        mask_local = df_f1[idx_local].astype(str).str.strip().str.upper().isin(locais)
        df_f2 = df_f1[mask_local]
        
        # 3. Status
        status_ok = ["SILVER", "GOLD", "DIAMOND"]
        mask_status = df_f2[idx_status].astype(str).str.strip().str.upper().isin(status_ok)
        df_f3 = df_f2[mask_status]
        
        # 4. Regra MG
        uf_str = df_f3[idx_uf].astype(str).str.strip().str.upper()
        st_str = df_f3[idx_status].astype(str).str.strip().str.upper()
        mask_mg = ~((uf_str == "MG") & (st_str == "SILVER"))
        df_f4 = df_f3[mask_mg]
        
        # 5. Transp
        transp_block = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
        mask_transp = ~df_f4[idx_transp].astype(str).str.strip().str.upper().isin(transp_block)
        df_final = df_f4[mask_transp]

        # Funil
        st.write("### 📉 Funil")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("1. Data", len(df_f1))
        c2.metric("2. Local", len(df_f2))
        c3.metric("3. Status", len(df_f3))
        c4.metric("4. MG", len(df_f4))
        c5.metric("5. Transp (Final)", len(df_final))

        if len(df_final) > 0:
            # Exportação
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
            cols_existentes = [c for c in cols_to_drop if c in df_final.columns]
            df_export = df_final.drop(columns=cols_existentes)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                wb = writer.book
                ws = writer.sheets['Sheet1']
                fmt_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                fmt_geral = wb.add_format({'border': 1})
                
                ws.conditional_format(0, 0, len(df_export)-1, len(df_export.columns)-1, 
                                    {'type': 'no_blanks', 'format': fmt_geral})
                
                if len(df_export.columns) > 5:
                    ws.set_column(5, 5, 15, fmt_moeda)
            
            st.download_button("📥 Baixar Planilha", output.getvalue(), "Cargas_Filtradas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        else:
            st.warning("Nenhum dado sobrou após os filtros.")
