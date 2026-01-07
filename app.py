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
    string_val = ""
    n += 1
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        string_val = chr(65 + remainder) + string_val
    return string_val

def carregar_dados_blindado(uploaded_file):
    """Tenta ler o arquivo de todas as formas possíveis."""
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
    for encoding in ['utf-8', 'latin-1', 'cp1252']:
        try:
            text = bytes_data.decode(encoding)
            
            # Tenta HTML (Tabela de site salva como excel)
            try:
                dfs = pd.read_html(io.StringIO(text), header=None)
                if dfs: return dfs[0]
            except:
                pass
            
            # Tenta CSV separado por Tabulação
            try:
                df = pd.read_csv(io.StringIO(text), sep='\t', header=None, engine='python')
                if df.shape[1] > 1: return df
            except:
                pass
        except:
            continue
            
    return None

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("⚙️ Configurações")
    
    st.info("⚠️ Atenção: O Python conta a partir do 0 (A=0, B=1). Use o 'Mapeador' na tela principal se tiver dúvida.")
    
    st.subheader("1. Data do Plantão")
    data_ref = st.date_input("Data de Início", datetime.now().date())
    
    st.divider()
    
    st.subheader("2. Mapeamento de Colunas")
    # VALORES JÁ CORRIGIDOS PARA O PYTHON (Subtraí 1 do que você usava)
    idx_data = st.number_input("DATA (L=11)", value=11, min_value=0, help="Coluna L no Excel é 11 aqui")
    idx_local = st.number_input("LOCAL (E=4)", value=4, min_value=0, help="Coluna E no Excel é 4 aqui")
    idx_uf = st.number_input("UF (I=8)", value=8, min_value=0, help="Coluna I no Excel é 8 aqui")
    idx_transp = st.number_input("TRANSP (K=10)", value=10, min_value=0, help="Coluna K no Excel é 10 aqui")
    idx_status = st.number_input("STATUS (P=15)", value=15, min_value=0, help="Coluna P no Excel é 15 aqui")

# --- CORPO PRINCIPAL ---
st.title("🚀 Processador de Cargas Pro")

uploaded_file = st.file_uploader("Arraste seu arquivo aqui", type=["xls", "xlsx", "xlsm", "csv", "txt"])

if uploaded_file:
    df_raw = carregar_dados_blindado(uploaded_file)

    if df_raw is None:
        st.error("❌ Não foi possível ler o arquivo. Ele pode estar corrompido ou em formato desconhecido.")
        st.stop()

    # --- MAPEADOR VISUAL (A SOLUÇÃO DO PROBLEMA) ---
    with st.expander("🕵️‍♀️ Mapeador de Colunas (Confira os números aqui)", expanded=True):
        st.write("Verifique se o número da coluna bate com os dados:")
        
        # Cria uma visualização das primeiras linhas
        preview = df_raw.head(3).T.reset_index()
        preview.columns = ["Índice Python", "Linha 1", "Linha 2", "Linha 3"]
        
        # Adiciona a Letra do Excel para facilitar
        preview.insert(1, "Letra Excel", [get_col_letter(i) for i in preview["Índice Python"]])
        
        st.dataframe(preview, height=300, use_container_width=True)

    st.divider()

    # --- PROCESSAMENTO ---
    try:
        # Tratamento de Data
        df_raw[idx_data] = pd.to_datetime(df_raw[idx_data], dayfirst=True, errors='coerce')
        df_limpo = df_raw.dropna(subset=[idx_data]).copy()
        
        # Confirmação Visual das Datas
        min_dt = df_limpo[idx_data].min()
        max_dt = df_limpo[idx_data].max()
        
        col_ok, col_info = st.columns([1, 3])
        if pd.isna(min_dt):
            st.error(f"❌ Nenhuma data encontrada na Coluna {idx_data} ({get_col_letter(idx_data)}). Verifique o mapeador acima!")
            st.stop()
        else:
            col_ok.success("✅ Coluna de Datas OK")
            col_info.info(f"O arquivo contém dados de **{min_dt.strftime('%d/%m %H:%M')}** até **{max_dt.strftime('%d/%m %H:%M')}**")

    except Exception as e:
        st.error(f"Erro ao ler datas: {e}")
        st.stop()

    # Filtros
    inicio = datetime.combine(data_ref, time(17, 0))
    fim = datetime.combine(data_ref + timedelta(days=1), time(7, 0))
    
    st.markdown(f"**Filtrando Plantão:** `{inicio}` até `{fim}`")

    # 1. Data
    df_f1 = df_limpo[(df_limpo[idx_data] >= inicio) & (df_limpo[idx_data] <= fim)]
    
    # 2. Local
    locais = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
    df_f2 = df_f1[df_f1[idx_local].astype(str).str.strip().str.upper().isin(locais)]
    
    # 3. Status
    status_ok = ["SILVER", "GOLD", "DIAMOND"]
    df_f3 = df_f2[df_f2[idx_status].astype(str).str.strip().str.upper().isin(status_ok)]
    
    # 4. Regra MG
    mask_mg = ~((df_f3[idx_uf].astype(str).str.strip().str.upper() == "MG") & 
                (df_f3[idx_status].astype(str).str.strip().str.upper() == "SILVER"))
    df_f4 = df_f3[mask_mg]
    
    # 5. Transp
    transp_block = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
    df_final = df_f4[~df_f4[idx_transp].astype(str).str.strip().str.upper().isin(transp_block)]

    # Funil
    st.write("### 📉 Funil de Resultados")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("1. Data", len(df_f1))
    c2.metric("2. Local", len(df_f2))
    c3.metric("3. Status", len(df_f3))
    c4.metric("4. MG", len(df_f4))
    c5.metric("5. Final", len(df_final))

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
            
            # Tenta formatar a coluna de valor (nova coluna F -> índice 5)
            if len(df_export.columns) > 5:
                ws.set_column(5, 5, 15, fmt_moeda)
        
        st.download_button("📥 Baixar Planilha Pronta", output.getvalue(), "Cargas_Filtradas.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    elif len(df_f1) == 0:
        st.warning("⚠️ O filtro de data removeu todas as linhas. Verifique se a 'Data de Início' na lateral corresponde ao arquivo.")
    else:
        st.warning("⚠️ Nenhum dado sobrou após aplicar todos os filtros.")
