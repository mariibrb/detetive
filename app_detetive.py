import streamlit as st
import pandas as pd
import io

def processar_cruzamento_fiscal(file_garimpeiro, file_sped):
    """
    Função principal que recebe os arquivos em memória, processa a leitura rigorosa
    das abas e colunas especificadas e retorna um buffer de Excel com o resultado.
    """
    try:
        # Leitura rigorosa do arquivo do Garimpeiro (Matriz)
        try:
            df_garimpeiro = pd.read_excel(file_garimpeiro, sheet_name="Geral_Filtrado", dtype=str)
        except ValueError:
            return None, "Erro: A aba 'Geral_Filtrado' não foi encontrada no arquivo do Garimpeiro."
        
        if "Chave" not in df_garimpeiro.columns:
            return None, "Erro: A coluna 'Chave' não foi encontrada na aba 'Geral_Filtrado'."

        # Leitura rigorosa do arquivo do SPED (Alvo da Auditoria)
        try:
            df_sped = pd.read_excel(file_sped, sheet_name="C100 - DOCUMENTOS", dtype=str)
        except ValueError:
            return None, "Erro: A aba 'C100 - DOCUMENTOS' não foi encontrada no arquivo do SPED."
        
        if "CHV_NFE" not in df_sped.columns:
            return None, "Erro: A coluna 'CHV_NFE' não foi encontrada na aba 'C100 - DOCUMENTOS'."

        # Higienização dos dados: dropna, conversão para string e remoção de espaços
        chaves_matriz = df_garimpeiro['Chave'].dropna().astype(str).str.strip()
        chaves_alvo = df_sped['CHV_NFE'].dropna().astype(str).str.strip()

        # Conversão para conjuntos (sets) para cruzamento de alta performance
        set_garimpeiro = set(chaves_matriz)
        set_sped = set(chaves_alvo)

        # Regras de Agregação da Análise
        # 1. Notas Faltando no SPED (Estão no Garimpeiro, mas não no SPED)
        faltando_no_sped = set_garimpeiro - set_sped
        
        # 2. Notas A Mais no SPED (Estão no SPED, mas não no Garimpeiro)
        a_mais_no_sped = set_sped - set_garimpeiro

        # Preparação dos DataFrames de saída
        df_faltando = pd.DataFrame(list(faltando_no_sped), columns=['Chaves_Faltando_SPED'])
        df_a_mais = pd.DataFrame(list(a_mais_no_sped), columns=['Chaves_Excedentes_SPED'])

        # Geração do arquivo Excel em memória (Buffer)
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_faltando.to_excel(writer, index=False, sheet_name='Faltando no SPED')
            df_a_mais.to_excel(writer, index=False, sheet_name='A mais no SPED')
            
            # Ajuste de largura das colunas para melhor visualização
            worksheet_faltando = writer.sheets['Faltando no SPED']
            worksheet_faltando.set_column('A:A', 50)
            
            worksheet_a_mais = writer.sheets['A mais no SPED']
            worksheet_a_mais.set_column('A:A', 50)

        output.seek(0)
        return output, "Sucesso"

    except Exception as e:
        return None, f"Erro inesperado durante o processamento: {str(e)}"

# Interface gráfica no navegador via Streamlit
st.set_page_config(page_title="Projeto Detetive - Auditoria", layout="wide")

st.title("Projeto Detetive: Cruzamento Fiscal (Garimpeiro x SPED)")
st.markdown("---")

st.markdown("""
**Instruções de Uso:**
1. Insira o relatório do **Garimpeiro** (Matriz). Deve conter a aba `Geral_Filtrado` com a coluna `Chave`.
2. Insira o relatório do **SPED**. Deve conter a aba `C100 - DOCUMENTOS` com a coluna `CHV_NFE`.
""")

col1, col2 = st.columns(2)

with col1:
    arquivo_garimpeiro = st.file_uploader("Upload Relatório Garimpeiro (.xlsx)", type=['xlsx', 'xls'], key="garimpeiro")

with col2:
    arquivo_sped = st.file_uploader("Upload Relatório SPED (.xlsx)", type=['xlsx', 'xls'], key="sped")

if arquivo_garimpeiro and arquivo_sped:
    if st.button("Executar Cruzamento Fiscal", type="primary"):
        with st.spinner("Processando dados e aplicando regras de matriz..."):
            buffer_excel, status_msg = processar_cruzamento_fiscal(arquivo_garimpeiro, arquivo_sped)
            
            if buffer_excel is not None:
                st.success("Cruzamento finalizado com sucesso! Faça o download do relatório abaixo.")
                st.download_button(
                    label="📥 Baixar Relatório de Discrepâncias (.xlsx)",
                    data=buffer_excel,
                    file_name="Relatorio_Auditoria_Detetive.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error(status_msg)
