import streamlit as st
import pandas as pd
import io

def processar_relatorio_detetive(file_garimpeiro, file_sped):
    try:
        # 1. Leitura das abas conforme sua estrutura
        df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        
        # Limpeza rigorosa das chaves para não dar erro de espaço
        df_garimpeiro['Chave'] = df_garimpeiro['Chave'].astype(str).str.strip()
        chaves_sped_set = set(df_sped['CHV_NFE'].dropna().astype(str).str.strip())
        chaves_garimpeiro_set = set(df_garimpeiro['Chave'].dropna().astype(str).str.strip())

        # 2. Criando a coluna de indicação (Ladainha resolvida: coluna criada no Garimpeiro)
        df_garimpeiro['CONSTA_NO_SPED'] = df_garimpeiro['Chave'].apply(
            lambda x: "SIM" if x in chaves_sped_set else "NÃO - NOTA FALTANDO NO SPED"
        )

        # 3. Pegando dados das notas que SÓ estão no SPED
        df_so_no_sped = df_sped[~df_sped['CHV_NFE'].astype(str).str.strip().isin(chaves_garimpeiro_set)].copy()
        
        # Ajustando colunas para o anexo final
        if not df_so_no_sped.empty:
            df_so_no_sped = df_so_no_sped.rename(columns={'CHV_NFE': 'Chave'})
            df_so_no_sped['CONSTA_NO_SPED'] = "ERRO - NOTA EXCEDENTE (NÃO ESTÁ NO GARIMPEIRO)"

        # 4. Unificando as bases (Garimpeiro em cima, sobras do SPED embaixo)
        df_final = pd.concat([df_garimpeiro, df_so_no_sped], ignore_index=True, sort=False)

        # 5. Gerando o Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_final.to_excel(writer, index=False, sheet_name='Detetive_Fiscal')
            
            # Ajuste de layout para as colunas não ficarem espremidas
            worksheet = writer.sheets['Detetive_Fiscal']
            for i, col in enumerate(df_final.columns):
                worksheet.set_column(i, i, 30)
            
        output.seek(0)
        return output, "Sucesso"
    except Exception as e:
        return None, f"Erro: {str(e)}"

# --- INTERFACE NO CHROME ---
st.set_page_config(page_title="Projeto Detetive", layout="wide")
st.title("🕵️ Projeto Detetive: Auditoria de Chaves")
st.markdown("---")

# RESTAURADO: Os dois botões de upload que você amou
col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Base Matriz")
    arquivo_g = st.file_uploader("Upload do Relatório Garimpeiro", type=['xlsx'], key="up_garimpeiro")

with col2:
    st.subheader("2. Base de Confronto")
    arquivo_s = st.file_uploader("Upload do Arquivo SPED", type=['xlsx'], key="up_sped")

# Só executa se os dois arquivos forem enviados
if arquivo_g and arquivo_s:
    st.markdown("---")
    if st.button("🚀 GERAR RELATÓRIO DE AUDITORIA", type="primary"):
        with st.spinner("O Detetive está analisando as chaves..."):
            resultado, mensagem = processar_relatorio_detetive(arquivo_g, arquivo_s)
            
            if resultado:
                st.success("Auditoria concluída! O relatório unificado está pronto.")
                st.download_button(
                    label="📥 Baixar Planilha de Divergências",
                    data=resultado,
                    file_name="Auditoria_Detetive_Final.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(mensagem)
