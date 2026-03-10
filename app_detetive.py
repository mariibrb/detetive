
import streamlit as st
import pandas as pd
import io

def processar_detetive_fiel(file_garimpeiro, file_sped):
    try:
        # 1. Leitura Íntegra (Matriz Garimpeiro e Alvo SPED)
        df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        
        # Higienização de Chaves
        chaves_sped_set = set(df_sped['CHV_NFE'].dropna().astype(str).str.strip())
        chaves_garimpeiro_set = set(df_garimpeiro['Chave'].dropna().astype(str).str.strip())

        # 2. CRIANDO A COLUNA NO GARIMPEIRO (O que você pediu)
        df_garimpeiro['CONSTA_NO_SPED'] = df_garimpeiro['Chave'].astype(str).str.strip().apply(
            lambda x: "SIM" if x in chaves_sped_set else "NÃO - NOTA FALTANDO NO SPED"
        )

        # 3. IDENTIFICANDO O QUE SÓ TEM NO SPED (Dados das notas excedentes)
        # Aqui pegamos as linhas do SPED que não estão no Garimpeiro
        df_excedentes_sped = df_sped[~df_sped['CHV_NFE'].astype(str).str.strip().isin(chaves_garimpeiro_set)].copy()
        
        # Renomeamos para manter o padrão da planilha única
        df_excedentes_sped = df_excedentes_sped.rename(columns={'CHV_NFE': 'Chave'})
        df_excedentes_sped['CONSTA_NO_SPED'] = "ERRO - NOTA SOBRANDO NO SPED (NÃO ESTÁ NO GARIMPEIRO)"

        # 4. UNIFICAÇÃO FINAL (Mantendo todas as colunas de ambos)
        # O concat preserva as colunas do Garimpeiro e adiciona as do SPED onde houver excedente
        df_final = pd.concat([df_garimpeiro, df_excedentes_sped], ignore_index=True, sort=False)

        # 5. Geração do Arquivo
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='DETETIVE_FINAL')
            
            # Ajuste de Layout
            worksheet = writer.sheets['DETETIVE_FINAL']
            for i, col in enumerate(df_final.columns):
                worksheet.set_column(i, i, 25)
            
        output.seek(0)
        return output, "Sucesso"
    except Exception as e:
        return None, f"Erro detalhado: {str(e)}"

# Interface
st.set_page_config(page_title="Projeto Detetive", layout="wide")
st.title("🕵️ Detetive: Auditoria com Colunas de Status")

g = st.file_uploader("Arquivo Garimpeiro", type=['xlsx'])
s = st.file_uploader("Arquivo SPED", type=['xlsx'])

if g and s:
    if st.button("GERAR AUDITORIA COMPLETA"):
        res, msg = processar_detetive_fiel(g, s)
        if res:
            st.success("Relatório unificado com as colunas criadas!")
            st.download_button("📥 Baixar Relatório Detetive.xlsx", data=res, file_name="Detetive_Final.xlsx")
        else:
            st.error(msg)
