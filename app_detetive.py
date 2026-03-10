import streamlit as st
import pandas as pd
import io

def processar_auditoria_detetive(file_garimpeiro, file_sped):
    try:
        # 1. Leitura das abas específicas conforme sua regra
        df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        
        # Limpeza de chaves
        df_garimpeiro['Chave'] = df_garimpeiro['Chave'].astype(str).str.strip()
        chaves_sped_set = set(df_sped['CHV_NFE'].dropna().astype(str).str.strip())
        chaves_garimpeiro_set = set(df_garimpeiro['Chave'].dropna().astype(str).str.strip())

        # 2. Criando a coluna de indicação no Garimpeiro
        df_garimpeiro['CONSTA_NO_SPED'] = df_garimpeiro['Chave'].apply(
            lambda x: "SIM" if x in chaves_sped_set else "NÃO - NOTA FALTANDO NO SPED"
        )

        # 3. Identificando notas que só existem no SPED
        df_so_no_sped = df_sped[~df_sped['CHV_NFE'].astype(str).str.strip().isin(chaves_garimpeiro_set)].copy()
        
        if not df_so_no_sped.empty:
            df_so_no_sped = df_so_no_sped.rename(columns={'CHV_NFE': 'Chave'})
            df_so_no_sped['CONSTA_NO_SPED'] = "ERRO - NOTA EXCEDENTE (SÓ NO SPED)"

        # 4. Unificação (Garimpeiro com coluna nova + Excedentes do SPED no final)
        df_final = pd.concat([df_garimpeiro, df_so_no_sped], ignore_index=True, sort=False)

        # 5. Geração do Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Auditoria_Detetive')
            worksheet = writer.sheets['Auditoria_Detetive']
            for i, col in enumerate(df_final.columns):
                worksheet.set_column(i, i, 30)
            
        output.seek(0)
        return output, "Sucesso"
    except Exception as e:
        return None, f"Erro: {str(e)}"

# --- INTERFACE VISUAL NO CHROME ---
st.set_page_config(page_title="Projeto Detetive", layout="wide")
st.title("🕵️ Projeto Detetive")

# Garantindo os DOIS campos de upload na tela
col1, col2 = st.columns(2)

with col1:
    st.subheader("Arquivo 1")
    arquivo_garimpeiro = st.file_uploader("Relatório Garimpeiro", type=['xlsx'], key="g_up")

with col2:
    st.subheader("Arquivo 2")
    arquivo_sped = st.file_uploader("Relatório SPED", type=['xlsx'], key="s_up")

# Apenas se ambos estiverem presentes, o botão aparece
if arquivo_garimpeiro and arquivo_sped:
    if st.button("EXECUTAR CRUZAMENTO", type="primary"):
        res, msg = processar_auditoria_detetive(arquivo_garimpeiro, arquivo_sped)
        if res:
            st.success("Relatório gerado!")
            st.download_button("Baixar Auditoria Unificada", data=res, file_name="Detetive_Fiscal_Final.xlsx")
        else:
            st.error(msg)
