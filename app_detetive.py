import streamlit as st
import pandas as pd
import io

def processar_cruzamento_unificado(file_garimpeiro, file_sped):
    try:
        # Leitura das abas específicas
        df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        
        # Higienização rigorosa
        chaves_matriz = set(df_garimpeiro['Chave'].dropna().astype(str).str.strip())
        chaves_alvo = set(df_sped['CHV_NFE'].dropna().astype(str).str.strip())

        # Cálculo das diferenças
        faltando = chaves_matriz - chaves_alvo
        excedente = chaves_alvo - chaves_matriz

        # Criação do relatório UNIFICADO
        lista_final = []
        for chave in faltando:
            lista_final.append({"Chave_Apurada": chave, "Status_Auditoria": "FALTANDO NO SPED"})
        for chave in excedente:
            lista_final.append({"Chave_Apurada": chave, "Status_Auditoria": "A MAIS NO SPED (EXCEDENTE)"})

        df_unificado = pd.DataFrame(lista_final)

        # Se não houver erros, cria um aviso
        if df_unificado.empty:
            df_unificado = pd.DataFrame([{"Aviso": "Parabéns! SPED e Garimpeiro estão 100% iguais."}])

        # Geração do arquivo Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_unificado.to_excel(writer, index=False, sheet_name='Resultado_Detetive')
            
            # Formatação visual automática
            worksheet = writer.sheets['Resultado_Detetive']
            worksheet.set_column('A:A', 50) # Largura da Chave
            worksheet.set_column('B:B', 30) # Largura do Status
            
        output.seek(0)
        return output, "Sucesso"
    except Exception as e:
        return None, f"Erro: {str(e)}"

# Interface Streamlit
st.set_page_config(page_title="Detetive - Unificado", layout="wide")
st.title("🕵️ Projeto Detetive: Relatório Unificado")

col1, col2 = st.columns(2)
with col1:
    arq_garimpeiro = st.file_uploader("Upload Garimpeiro (Geral_Filtrado)", type=['xlsx'])
with col2:
    arq_sped = st.file_uploader("Upload SPED (C100 - DOCUMENTOS)", type=['xlsx'])

if arq_garimpeiro and arq_sped:
    if st.button("Gerar Relatório Unificado", type="primary"):
        with st.spinner("Cruzando dados..."):
            res, msg = processar_cruzamento_unificado(arq_garimpeiro, arq_sped)
            if res:
                st.success("Análise concluída em uma única aba!")
                st.download_button("📥 Baixar Relatório Único (.xlsx)", data=res, file_name="Detetive_Unificado.xlsx")
            else:
                st.error(msg)
