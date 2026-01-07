import pandas as pd
import streamlit as st
import io
import string
from datetime import datetime, timedelta, time

# --- CONFIGURAÇÃO DA PÁGINA ---
st.set_page_config(page_title="Processador de Cargas Pro", page_icon="🚀", layout="wide")

# --- FUNÇÕES DE ENGENHARIA ---

def get_col_letter(n):
    """Converte 0->A, 1->B para ajudar o usuário visualmente."""
    try:
        string_val = ""
        n = int(n) + 1
        while n > 0:
            n, remainder = divmod(n - 1, 26)
            string_val = chr(65 + remainder) + string_val
        return string_val
    except:
        return "?"

def carregar_dados_blindado(uploaded_file):
    """Tenta ler o arquivo de todas as formas possíveis, incluindo UTF-16."""
    bytes_data = uploaded_file.getvalue()
    
    # 1. Tentar Excel Padrão (.xlsx)
    try:
        return pd.read_excel(io.BytesIO(bytes_data), header=None)
    except:
        pass
        
    # 2. Tentar Excel Antigo (.xls) com engine 'xlrd'
    try:
        return pd.read_excel(io.BytesIO(bytes_data), header=None, engine='xlrd')
    except:
        pass

    # 3. Tentar HTML ou Texto (Sistemas Legados)
    # ADICIONEI 'utf-16' AQUI PARA RESOLVER SEU PROBLEMA
    encodings_to_try = ['utf-16', 'utf-8', 'latin-1', 'cp1252']
    
    for encoding in encodings_to_try:
        try:
            # Tenta decodificar o texto
            text = bytes_data.decode(encoding)
            
            # Tenta ler como CSV separado por Tabulação (MUITO COMUM EM UTF-16)
            try:
                df = pd.read_csv(io.StringIO(text), sep='\t', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass

            # Tenta ler como HTML
            try:
                dfs = pd.read_html(io.StringIO(text), header=None)
                if dfs: return dfs[0]
            except:
                pass
            
            # Tenta ler como CSV separado por Ponto e Vírgula
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
    
    st.info("⚠️ O código já está configurado com os índices corretos (A=0). Só mude se a planilha mudar.")
    
    st.subheader("1. Data do Plantão")
    data_ref = st.date_input("Data de Início", datetime.now().date())
    
    st.divider()
    
    st.subheader("2. Mapeamento de Colunas")
    # VALORES PADRÃO CORRIGIDOS PARA PYTHON
    idx_data = st.number_input("DATA (L=11)", value=11, min_value=0, help="No Python, L é 11")
    idx_local = st.number_input("LOCAL (E=4)", value=4, min_value=0, help="No Python, E é 4")
    idx_uf = st.number_input("UF (I=8)", value=8, min_value=0, help="No Python, I é 8")
    idx_transp = st.number_input("TRANSP (K=10)", value=10, min_value=0, help="No Python, K é 10")
    idx_status = st.number_input("STATUS (P=15)", value=15, min_value=0, help="No Python, P é 15")

# --- CORPO PRINCIPAL ---
st.title("🚀 Processador de Cargas Pro")

uploaded_file = st.file_uploader("Arraste seu arquivo aqui", type=["xls", "xlsx", "xlsm", "csv", "txt"])

if uploaded_file:
    df_raw = carregar_dados_blindado(uploaded_file)

    if df_raw is None:
        st.error("❌ Não foi possível ler o arquivo. Ele parece estar em um formato binário desconhecido.")
        st.stop()

    # --- MAPEADOR VISUAL (AGORA COM LETRAS) ---
    with st.expander("🕵️‍♀️ Visualizador de Colunas (Dados Reais)", expanded=True):
        st.write("Confira abaixo se os dados aparecem corretamente:")
        
        # Cria visualização
        preview = df_raw.head(3).T.reset_index()
        preview.columns = ["Índice Python", "Linha 1", "Linha 2", "Linha 3"]
        
        # Adiciona Letra do Excel
        preview.insert(1, "Letra Excel", [get_col_letter(i) for i in preview["Índice Python"]])
        
        st.dataframe(preview, height=300, use_container_width=True)

    st.divider()

    # --- PROCESSAMENTO ---
    try:
        # 1. Tratamento de Data
        # Força conversão de datas
        df_raw[idx_data] = pd.to_datetime(df_raw[idx_data], dayfirst=True, errors='coerce')
        
        # Remove linhas onde a data é inválida (cabeçalhos ou vazios)
        df_limpo = df_raw.dropna(subset=[idx_data]).copy()
        
        if len(df_limpo) == 0:
            st.warning("⚠️ Nenhuma data válida encontrada na coluna indicada. Verifique o Mapeador acima.")
            st.stop()
            
        # Confirmação Visual
        min_dt = df_limpo[idx_data].min()
        max_dt = df_limpo[idx_data].max()
        
        col_ok, col_info = st.columns([1, 3])
        col_ok.success("✅ Datas Lidas!")
        col_info.info(f"Período no arquivo: **{min_dt.strftime('%d/%m %H:%M')}** até **{max_dt.strftime('%d/%m %H:%M')}**")

        # 2. Definição do Filtro
        inicio = datetime.combine(data_ref, time(17, 0))
        fim = datetime.combine(data_ref + timedelta(days=1), time(7, 0))
        
        st.markdown(f"**🔎 Filtrando por:** `{inicio}` até `{fim}`")

        # 3. Aplicando Filtros Sequencialmente
        
        # Data
        df_f1 = df_limpo[(df_limpo[idx_data] >= inicio) & (df_limpo[idx_data] <= fim)]
        
        # Local
        locais = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
        # Converte para string e maiúsculo para garantir
        df_f2 = df_f1[df_f1[idx_local].astype(str).str.strip().str.upper().isin(locais)]
        
        # Status
        status_ok = ["SILVER", "GOLD", "DIAMOND"]
        df_f3 = df_f2[df_f2[idx_status].astype(str).str.strip().str.upper().isin(status_ok)]
        
        # Regra MG + Silver
        uf_col = df_f3[idx_uf].astype(str).str.strip().str.upper()
        st_col = df_f3[idx_status].astype(str).str.strip().str.upper()
        # Mantém o que NÃO FOR (MG E SILVER)
        df_f4 = df_f3[~((uf_col == "MG") & (st_col == "SILVER"))]
        
        # Transportadoras
        transp_block = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
        df_final = df_f4[~df_f4[idx_transp].astype(str).str.strip().str.upper().isin(transp_block)]

        # --- FUNIL DE RESULTADOS ---
        st.write("### 📉 Funil de Resultados")
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("1. Data", len(df_f1))
        c2.metric("2. Local", len(df_f2))
        c3.metric("3. Status", len(df_f3))
        c4.metric("4. MG", len(df_f4))
        c5.metric("5. Final", len(df_final))

        # --- EXPORTAÇÃO ---
        if len(df_final) > 0:
            # Remove colunas desnecessárias
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
            # Filtra apenas as que existem
            cols_existentes = [c for c in cols_to_drop if c in df_final.columns]
            df_export = df_final.drop(columns=cols_existentes)
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_export.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                wb = writer.book
                ws = writer.sheets['Sheet1']
                fmt_moeda = wb.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                fmt_geral = wb.add_format({'border': 1})
                
                # Formata tudo
                ws.conditional_format(0, 0, len(df_export)-1, len(df_export.columns)-1, 
                                    {'type': 'no_blanks', 'format': fmt_geral})
                
                # Tenta formatar a coluna de valor (nova coluna F -> índice 5)
                if len(df_export.columns) > 5:
                    ws.set_column(5, 5, 15, fmt_moeda)
            
            st.success(f"✅ Sucesso! {len(df_final)} linhas geradas.")
            st.download_button("📥 Baixar Planilha Pronta", output.getvalue(), "Cargas_Filtradas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
        
        elif len(df_f1) == 0:
            st.warning(f"⚠️ O filtro de data removeu tudo. Verifique a Data de Início ({data_ref}).")
        else:
            st.warning("⚠️ Nenhum dado sobrou após os filtros.")

    except Exception as e:
        st.error(f"Erro durante o processamento: {e}")
