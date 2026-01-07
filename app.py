import pandas as pd
from datetime import datetime, timedelta, time
import streamlit as st
import io

# Configuração da Página
st.set_page_config(page_title="Macro Cargas", page_icon="🚛", layout="centered")

# Título da Aplicação
st.title("🚛 Processador de Cargas - Critério Melhorado")
st.markdown("Faça upload da planilha para filtrar as cargas e formatar automaticamente.")

# 1. Upload do Arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "xlsm"])

if uploaded_file is not None:
    df = None
    # --- BLOCO DE LEITURA BLINDADA (FINAL) ---
    try:
        # TENTATIVA 1: Leitura Padrão (Excel moderno .xlsx)
        df = pd.read_excel(uploaded_file, header=None)
    except Exception:
        try:
            # TENTATIVA 2: Forçar engine 'xlrd' para arquivos .xls antigos
            uploaded_file.seek(0) 
            df = pd.read_excel(uploaded_file, header=None, engine='xlrd')
        except Exception:
            # TENTATIVA 3: Arquivo HTML/XML "disfarçado"
            # Tenta UTF-8 primeiro, depois Latin-1 (para corrigir o UnicodeDecodeError)
            try:
                uploaded_file.seek(0)
                dfs_html = pd.read_html(uploaded_file, header=None)
            except UnicodeDecodeError:
                # CORREÇÃO DO ERRO DE CARACTERES ESTRANHOS
                uploaded_file.seek(0)
                dfs_html = pd.read_html(uploaded_file, header=None, encoding='latin-1')
            except Exception:
                # Última chance: cp1252 (outro padrão comum no Windows)
                try:
                    uploaded_file.seek(0)
                    dfs_html = pd.read_html(uploaded_file, header=None, encoding='cp1252')
                except:
                    dfs_html = None

            if dfs_html:
                df = dfs_html[0]
            else:
                st.error("Erro fatal: O arquivo não é um Excel válido e nem uma tabela HTML legível.")
                st.stop()
    
    # Se chegou aqui, o df existe. Vamos processar.
    if df is not None:
        try:
            # --- Definição do Plantão ---
            hoje = datetime.now().date()
            # Ajuste aqui se precisar mudar o horário do plantão
            inicio_plantao = datetime.combine(hoje, time(17, 0))
            fim_plantao = datetime.combine(hoje + timedelta(days=1), time(7, 0))

            st.info(f"Processando {len(df)} linhas originais. Plantão: {inicio_plantao} até {fim_plantao}")

            # --- Filtragem ---
            # Índices baseados na macro original:
            # Col L (Data) = 11 | Col E (Local) = 4 | Col P (Status) = 15 
            # Col I (UF) = 8    | Col K (Transp) = 10
            
            # 1. Datas (Coluna L / Index 11)
            df[11] = pd.to_datetime(df[11], errors='coerce')
            mascara_data = (df[11] >= inicio_plantao) & (df[11] <= fim_plantao)
            
            # 2. Local (Coluna E / Index 4)
            locais_validos = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
            mascara_local = df[4].astype(str).str.strip().isin(locais_validos)
            
            # 3. Status (Coluna P / Index 15)
            status_validos = ["SILVER", "GOLD", "DIAMOND"]
            mascara_status = df[15].astype(str).str.strip().str.upper().isin(status_validos)
            
            # 4. Exceção MG + SILVER (Coluna I / Index 8 e P / Index 15)
            # Excluir se for MG E Silver (usa ~ para inverter e manter o resto)
            mascara_mg_silver = ~((df[8].astype(str).str.strip().str.upper() == "MG") & 
                                  (df[15].astype(str).str.strip().str.upper() == "SILVER"))
            
            # 5. Transportadoras Bloqueadas (Coluna K / Index 10)
            transp_bloqueadas = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
            mascara_transp = ~df[10].astype(str).str.strip().str.upper().isin(transp_bloqueadas)

            # APLICAR TODOS OS FILTROS
            df_filtrado = df[mascara_data & mascara_local & mascara_status & mascara_mg_silver & mascara_transp].copy()

            # --- Exclusão de Colunas ---
            # Indices para remover: V(21), U(20)... A(0)
            cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
            
            # Remove apenas colunas que existem no dataframe
            cols_existentes = [c for c in cols_to_drop if c in df_filtrado.columns]
            df_final = df_filtrado.drop(columns=cols_existentes)

            # --- Exportação e Download ---
            output = io.BytesIO()
            
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df_final.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # Formatos
                formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
                formato_bordas = workbook.add_format({'border': 1})
                
                # Aplica bordas em tudo se houver dados
                if len(df_final) > 0:
                    worksheet.conditional_format(0, 0, len(df_final)-1, len(df_final.columns)-1, 
                                                {'type': 'no_blanks', 'format': formato_bordas})

                # Formato de Moeda na nova coluna F (Index 5)
                # A Coluna O (15) original vira a F (5) após os cortes
                if len(df_final.columns) > 5:
                    worksheet.set_column(5, 5, 15, formato_moeda) 
                
                # Ajuste de largura geral
                worksheet.set_column(0, 4, 12)
                worksheet.set_column(6, 20, 12)
            
            st.success(f"Sucesso! Linhas restantes: {len(df_final)}")
            
            st.download_button(
                label="📥 Baixar Arquivo Filtrado",
                data=output.getvalue(),
                file_name="Relatorio_Filtrado.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        except Exception as e:
            st.error(f"Erro durante o processamento dos dados: {e}")
            st.write("Dica: Verifique se as colunas da planilha mudaram de posição.")
