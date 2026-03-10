import streamlit as st
import pandas as pd
import io

def processar_auditoria_unificada(file_garimpeiro, file_sped):
    try:
        # 1. Leitura rigorosa dos dados
        df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        
        # Limpeza de espaços e nulos para comparação
        df_garimpeiro['Chave'] = df_garimpeiro['Chave'].astype(str).str.strip()
        chaves_sped = set(df_sped['CHV_NFE'].dropna().astype(str).str.strip())
        chaves_garimpeiro = set(df_garimpeiro['Chave'].dropna().astype(str).str.strip())

        # 2. Criar a coluna indicadora no Garimpeiro
        df_garimpeiro['STATUS_NO_SPED'] = df_garimpeiro['Chave'].apply(
            lambda x: "OK - CONSTA NO SPED" if x in chaves_sped else "ERRO - FALTANDO NO SPED"
        )

        # 3. Identificar o que só tem no SPED (Excedentes)
        so_no_sped = chaves_sped - chaves_garimpeiro
        df_excedentes = pd.DataFrame(list(so_no_sped), columns=['Chave'])
        df_excedentes['STATUS_NO_SPED'] = "ERRO - NOTA EXCEDENTE (SÓ NO SPED)"

        # 4. Unificar tudo (Garimpeiro + Excedentes do SPED)
        df_final = pd.concat([df_garimpeiro, df_excedentes], ignore_index=True)

        # 5. Gerar o Excel com formatação básica
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Auditoria_Unificada')
            worksheet = writer.sheets['Auditoria_Unificada']
            worksheet.set_column('A:A', 50)
            worksheet.set_column(0, df_final.shape[1], 20) # Ajusta largura de todas as colunas
            
        output.seek(0)
        return output, "Sucesso"
    except Exception as e:
        return None, f"Erro: {str(e)}"

st.set_page_config(page_title="Detetive Fiscal", layout="wide")
st.title("🕵️ Detetive: Garimpeiro x SPED (Aba Única)")

garimpeiro = st.file_uploader("Arquivo Garimpeiro", type=['xlsx'])
sped = st.file_uploader("Arquivo SPED", type=['xlsx'])

if garimpeiro and sped:
    if st.button("Gerar Auditoria Unificada"):
        res, msg = processar_auditoria_unificada(garimpeiro, sped)
        if res:
            st.success("Relatório gerado com sucesso!")
            st.download_button("Baixar Relatório Unificado", data=res, file_name="Auditoria_Final.xlsx")
        else:
            st.error(msg)
