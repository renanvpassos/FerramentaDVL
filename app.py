import streamlit as st
import pandas as pd
from io import BytesIO

# Configuração da página do Streamlit
st.set_page_config(page_title="Tratamento planilha Delaval", layout="centered")

st.title("📊 Tratamento planilha Delaval")
st.markdown("Insira os parâmetros abaixo, suba os arquivos e realize os cálculos automaticamente.")

# --- SEÇÃO 1: PARÂMETROS DA INTERFACE ---
st.subheader("1. Configurações Padrão")
# Campo do ICMS posicionado no topo do fluxo principal
valor_icms = st.text_input("Valor do ICMS padrão (será aplicado em todas as linhas):", value="18")

st.markdown("---")

# --- SEÇÃO 2: UPLOAD DOS ARQUIVOS ---
st.subheader("2. Upload dos Arquivos")
arquivo_bd = st.file_uploader("Suba a planilha de BANCO DE DADOS (Excel)", type=["xlsx", "xls"])
arquivo_tratar = st.file_uploader("Suba a PLANILHA A SER TRATADA (Excel)", type=["xlsx", "xls"])

if arquivo_bd and arquivo_tratar:
    st.success("Arquivos carregados com sucesso! Processando...")

    try:
        # Lendo as planilhas
        df_bd = pd.read_excel(arquivo_bd)
        df_tratar = pd.read_excel(arquivo_tratar)

        # Padronizando o nome das colunas para evitar erros de espaços ou maiúsculas/minúsculas
        df_bd.columns = df_bd.columns.str.strip()
        df_tratar.columns = df_tratar.columns.str.strip()

        # Validação de colunas necessárias no Banco de Dados
        colunas_bd_obrigatorias = ['Material', 'peso liquido']
        validacao_bd = all(col in df_bd.columns for col in colunas_bd_obrigatorias)

        # Validação de colunas necessárias na Planilha a ser tratada
        colunas_tratar_obrigatorias = ['PARTNUMBER', 'QUANTIDADE']
        validacao_tratar = all(col in df_tratar.columns for col in colunas_tratar_obrigatorias)

        if not validacao_bd:
            st.error(
                f"O Banco de Dados precisa conter as colunas: {colunas_bd_obrigatorias}. Encontradas: {list(df_bd.columns)}")
        elif not validacao_tratar:
            st.error(
                f"A Planilha a ser tratada precisa conter as colunas: {colunas_tratar_obrigatorias}. Encontradas: {list(df_tratar.columns)}")
        else:
            # --- Início do Tratamento ---

            # 1. Cria a coluna ICMS no final com o valor informado no campo acima
            df_tratar['ICMS'] = valor_icms

            # 2. Faz o cruzamento (Merge) entre as tabelas
            # Remove duplicadas do banco de dados na coluna chave para não duplicar linhas na planilha final
            df_bd_limpo = df_bd.drop_duplicates(subset=['Material'])

            df_resultado = pd.merge(
                df_tratar,
                df_bd_limpo[['Material', 'peso liquido']],
                left_on='PARTNUMBER',
                right_on='Material',
                how='left'
            )

            # 3. Converte colunas para numérico para evitar erros de tipo no cálculo
            df_resultado['QUANTIDADE'] = pd.to_numeric(df_resultado['QUANTIDADE'], errors='coerce').fillna(0)
            df_resultado['peso liquido'] = pd.to_numeric(df_resultado['peso liquido'], errors='coerce').fillna(0)

            # 4. Calcula o PESOTOTAL (criado logo após a coluna ICMS)
            df_resultado['PESOTOTAL'] = df_resultado['peso liquido'] * df_resultado['QUANTIDADE']

            # 5. Remove colunas extras trazidas do banco de dados para manter o layout limpo
            if 'Material' in df_resultado.columns:
                df_resultado = df_resultado.drop(columns=['Material', 'peso liquido'])

            # Exibe uma prévia do resultado na tela
            st.subheader("💡 Prévia do Resultado")
            st.dataframe(df_resultado.head(10))

            # --- Preparação do arquivo para Download ---
            saida_excel = BytesIO()
            with pd.ExcelWriter(saida_excel, engine='xlsxwriter') as writer:
                df_resultado.to_excel(writer, index=False, sheet_name='Planilha Tratada')
            saida_excel.seek(0)

            st.subheader("3. Baixar Arquivo Final")
            st.download_button(
                label="📥 Baixar Planilha Tratada (Excel)",
                data=saida_excel,
                file_name="planilha_tratada_final.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    except Exception as e:
        st.error(f"Ocorreu um erro ao processar os arquivos: {e}")
