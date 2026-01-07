import pandas as pd
from datetime import datetime, timedelta, time
import streamlit as st
import io

# Título da Aplicação Web
st.title("Processador de Cargas - Critério Melhorado")

# 1. Upload do Arquivo
uploaded_file = st.file_uploader("Selecione o arquivo Excel", type=["xlsx", "xls", "xlsm"])

if uploaded_file is not None:
    # Ler o arquivo
    df = pd.read_excel(uploaded_file, header=None) # header=None para usar índices numéricos (0=A, 1=B, etc)
    
    # --- Definição do Plantão ---
    hoje = datetime.now().date()
    # Lógica simplificada de datas (pode ser ajustada conforme fuso horário do servidor)
    inicio_plantao = datetime.combine(hoje, time(17, 0))
    fim_plantao = datetime.combine(hoje + timedelta(days=1), time(7, 0))

    # --- Filtragem ---
    # Coluna L é índice 11, E é 4, P é 15, I é 8, K é 10 (Python começa no 0)
    
    # 1. Datas (Coluna L / Index 11)
    df[11] = pd.to_datetime(df[11], errors='coerce')
    mascara_data = (df[11] >= inicio_plantao) & (df[11] <= fim_plantao)
    
    # 2. Local (Coluna E / Index 4) - "CD POUSO ALEGRE" e "POUSO ALEGRE HPC"
    locais_validos = ["CD POUSO ALEGRE", "POUSO ALEGRE HPC"]
    mascara_local = df[4].astype(str).str.strip().isin(locais_validos)
    
    # 3. Status (Coluna P / Index 15) - SILVER, GOLD, DIAMOND
    status_validos = ["SILVER", "GOLD", "DIAMOND"]
    mascara_status = df[15].astype(str).str.strip().str.upper().isin(status_validos)
    
    # 4. Exceção MG + SILVER (Coluna I / Index 8 e P / Index 15)
    # Excluir se for MG E Silver
    mascara_mg_silver = ~((df[8].astype(str).str.strip().str.upper() == "MG") & 
                          (df[15].astype(str).str.strip().str.upper() == "SILVER"))
    
    # 5. Transportadoras Bloqueadas (Coluna K / Index 10)
    transp_bloqueadas = ["JSL S A", "TRANSANTA RITA LTDA", "T G LOGISTICA E TRANSPORTES LTDA", "TRANSANTA RITA TRANSPORTES LTDA"]
    mascara_transp = ~df[10].astype(str).str.strip().str.upper().isin(transp_bloqueadas)

    # APLICAR TODOS OS FILTROS JUNTOS
    df_filtrado = df[mascara_data & mascara_local & mascara_status & mascara_mg_silver & mascara_transp].copy()

    # --- Exclusão de Colunas ---
    # Mapeamento do VBA: V, U, T, S, R, Q, N, M, J, G, F, E, D, C, A
    # Índices correspondentes: 21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0
    cols_to_drop = [21, 20, 19, 18, 17, 16, 13, 12, 9, 6, 5, 4, 3, 2, 0]
    
    # Remove colunas que existem no dataframe
    cols_existentes = [c for c in cols_to_drop if c in df_filtrado.columns]
    df_final = df_filtrado.drop(columns=cols_existentes)

    # --- Exportação e Download ---
    output = io.BytesIO()
    # Usando XlsxWriter para formatar moeda e bordas
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final.to_excel(writer, index=False, header=False, sheet_name='Sheet1')
        workbook = writer.book
        worksheet = writer.sheets['Sheet1']
        
        # Formato de Moeda (R$) na nova coluna F (Index 5 no Python, mas no Excel final verifique a posição)
        formato_moeda = workbook.add_format({'num_format': 'R$ #,##0.00', 'border': 1})
        formato_geral = workbook.add_format({'border': 1})
        
        # Aplicar formatação (exemplo genérico)
        worksheet.set_column('F:F', None, formato_moeda) 
    
    st.success(f"Linhas restantes: {len(df_final)}")
    
    st.download_button(
        label="Baixar Arquivo Filtrado",
        data=output.getvalue(),
        file_name="Relatorio_Filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
