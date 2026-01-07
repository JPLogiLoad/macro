import pandas as pd
from datetime import datetime, timedelta, time
import streamlit as st
import io

# Configuração da Página
st.set_page_config(page_title="Macro Cargas", page_icon="🚛", layout="centered")

st.title("🚛 Processador de Cargas - Critério Melhorado")
st.markdown("Faça upload da planilha para filtrar as cargas e formatar automaticamente.")

# --- SELETOR DE DATA (NOVIDADE) ---
# Permite escolher a data base do plantão (padrão é hoje)
col1, col2 = st.columns(2)
with col1:
    data_escolhida = st.date_input("Data de Início do Plantão", datetime.now().date())
with col2:
    st.write("") # Espaçamento

# 1. Upload do Arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "xlsm", "csv", "txt"])

if uploaded_file is not None:
    df = None
    
    # --- BLOCO DE LEITURA BLINDADA ---
    try:
        # TENTATIVA 1: Leitura Padrão
        df = pd.read_excel(uploaded_file, header=None)
    except Exception:
        try:
            # TENTATIVA 2: Excel antigo (xlrd)
            uploaded_file.seek(0) 
            df = pd.read_excel(uploaded_file, header=None, engine='xlrd')
        except Exception:
            try:
                # TENTATIVA 3: HTML/XML disfarçado
                uploaded_file.seek(0)
                bytes_data = uploaded_file.getvalue()
                
                html_text = None
                # Tenta decodificar manualmente
                for encoding in ['utf-8', 'latin-1', 'cp1252']:
                    try:
                        html_text = bytes_data.decode(encoding)
                        break 
                    except UnicodeDecodeError:
                        continue
                
                if html_text:
                    dfs_html = pd.read_html(io.StringIO(html_text), header=None)
                    if dfs_html:
                        df = dfs_html[0]
                    else:
                        raise Exception("HTML sem tabelas")
            
            except Exception:
                # TENTATIVA 4: Texto/CSV
                try:
                    uploaded_file.seek(0)
                    df = pd.read_csv(uploaded_file, sep='\t', header=None, encoding='latin-1', engine='python')
                    if df.shape[1] < 2:
                        uploaded_file.seek(0)
                        df = pd.read_csv(uploaded_file, sep=';', header=None, encoding='latin-1', engine='python')
                except Exception as e:
                    st.error(f"Erro fatal de leitura: {e}")
                    st.stop()

    if df is not None:
        try:
            # --- Definição do Plantão (USANDO A DATA ESCOLHIDA) ---
            # Inicia às 17:00 da data escolhida e vai até 07:00 do dia seguinte
            inicio_plantao = datetime.combine(data_escolhida, time(17, 0))
            fim_plantao = datetime.combine(data_escolhida + timedelta(days=1), time(7, 0))

            st.info(f"Filtro aplicado: Cargas entre **{inicio_plantao.strftime('%d/%m %H:%M')}** e **{fim_plantao.strftime('%d/%m %H:%M')}**")

            # Verifica colunas
            if df.shape[1] < 16:
                st.error(f"Erro: A planilha tem apenas {df.shape[1]} colunas (precisa de pelo menos 16).")
                st.stop()

            # --- FILTRAGEM ---
            
            # 1. Datas (Coluna L / Index 11)
            # CORREÇÃO AQUI: dayfirst=True para formato BR (Dia/Mês/Ano)
            df[11] = pd.to_datetime(df[11], dayfirst=True, errors='coerce')
            
            # Remove linhas onde a data não pôde ser lida (NaT)
            linhas_data_invalida = df[11].isna().sum()
            if linhas_data_invalida > 0:
                st.warning(f"Atenção: {linhas_data_invalida} linhas tinham datas inválidas e foram ignoradas.")

            mascara_data = (df[11] >= inicio_plantao) & (df[11] <= fim_plantao)
            
            # 2. Local (Coluna E / Index 4)
            locais_validos = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
            mascara_local = df[4].astype(str).str.strip().isin(locais_validos)
            
            # 3. Status (Coluna P / Index 15)
            status_validos = ["SILVER", "GOLD", "DIAMOND"]
            mascara_status = df[15].astype(str).str.strip().str.upper().isin(status_validos)
            
            # 4. Exceção MG + SILVER (Coluna I / Index 8 e P / Index 15)
            mascara_mg_silver = ~((df[8].astype(str).str.strip().str.upper() == "MG") & 
                                  (df[15].astype(str).str.strip().str.upper() == "SILVER"))
            
            # 5. Transportadoras Bloqueadas (Coluna K / Index 10)
            transp_bloqueadas = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
            mascara_transp = ~df[10].astype(str).str.strip().str.upper().isin(transp_bloqueadas)

            # Aplicar Filtros
            df_filtrado = df[mascara_data & mascara_local & mascara_status & mascara_mg_silver & mascara_transp].copy()

            # --- PROCESSAMENTO FINAL ---
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
            cols_existentes = [c for c in cols_to_drop if c in df_filtrado.columns]
            df_final = df_filtrado.drop(columns=cols_existentes)

            # --- EXPORTAÇÃO ---
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                formato_bordas = workbook.add_format({'border': 1})
                
                if len(df_final) > 0:
                    worksheet.conditional_format(0, 0, len(df_final)-1, len(df_final.columns)-1, 
                                                {'type': 'no_blanks', 'format': formato_bordas})

                if len(df_final.columns) > 5:
                    worksheet.set_column(5, 5, 15, formato_moeda) 
                
                worksheet.set_column(0, 4, 12)
                worksheet.set_column(6, 20, 12)
            
            if len(df_final) == 0:
                st.warning("O filtro resultou em 0 linhas! Verifique se a data escolhida no topo corresponde às datas do arquivo.")
            else:
                st.success(f"Sucesso! {len(df_final)} cargas encontradas.")
            
            st.download_button(
                label="📥 Baixar Arquivo Filtrado",
                data=output.getvalue(),
                file_name="Relatorio_Filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro no processamento: {e}")
