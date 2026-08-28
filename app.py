import streamlit as st
import pandas as pd
from pypdf import PdfReader
import io
import re
import os

# --- CONFIGURAÇÃO DAS EXPRESSÕES REGULARES ---
# ATENÇÃO: Você deve ajustar estas expressões para corresponder ao formato exato dos seus PDFs.
# As expressões abaixo são exemplos e precisarão ser customizadas.

def extract_data(pdf_file_bytes):
    """
    Função que lê o PDF e tenta extrair as informações usando regex.
    Esta função é o coração da lógica de extração.
    """
    try:
        reader = PdfReader(io.BytesIO(pdf_file_bytes))
        text = ""
        # Junta todo o texto do PDF em uma única string
        for page in reader.pages:
            text += page.extract_text() + "\n"
        
        # --- LÓGICA DE EXTRAÇÃO BASEADA EM PADRÕES ---
        
        # Exemplo de como procurar (você deve customizar estes padrões):
        
        # 1. Invoice (Material)
        # Tenta encontrar algo como "Invoice Number: [número]" ou similar
        invoice_match = re.search(r"Invoice\s*Number[:\s]*(\w+)", text, re.IGNORECASE)
        invoice_number = invoice_match.group(1).strip() if invoice_match else "NÃO ENCONTRADO"

        # 2. PartNumber
        # Tenta encontrar números de 9 dígitos (ex: PartNumber: 123456789)
        partnumber_matches = re.findall(r"PartNumber[:\s]*(\d{9})", text, re.IGNORECASE)
        partnumbers = [match[0] for match in partnumber_matches] if partnumber_matches else ["NÃO ENCONTRADO"]

        # 3. Quantidade (Qty/Unit)
        # Tenta encontrar números decimais ou inteiros após uma descrição de item
        # Exemplo: Procurar por "Quantity" seguido de um número
        quantity_match = re.search(r"Quantity[:\s]*([\d\.]+)", text, re.IGNORECASE)
        quantity = quantity_match.group(1).strip() if quantity_match else "NÃO ENCONTRADO"

        # 4. PesoTotal (Net)
        # Tenta encontrar valores monetários ou pesos (com casas decimais)
        # Exemplo: Procurar por "Peso Total: 15.50 kg"
        weight_match = re.search(r"PesoTotal[:\s]*([\d\.]+)", text, re.IGNORECASE)
        weight_total = weight_match.group(1).strip() if weight_match else "NÃO ENCONTRADO"
        
        # 5. PrecoTotal (Total)
        # Exemplo: Procurar por "Preco Total: 1234.56"
        price_match = re.search(r"PrecoTotal[:\s]*([\d\.]+)", text, re.IGNORECASE)
        price_total = price_match.group(1).strip() if price_match else "NÃO ENCONTRADO"
        
        
        # Estruturação dos dados
        extracted_data = {
            "Invoice": invoice_number,
            "PartNumber": "; ".join(partnumbers), # Junta vários se encontrado
            "Quantidade": quantity,
            "PesoTotal": weight_total,
            "PrecoTotal": price_total
        }
        
        return extracted_data

    except Exception as e:
        st.error(f"Ocorreu um erro durante a extração: {e}")
        return None

# --- INTERFACE STREAMLIT ---

st.title("📄 Extrator de Dados de PDF para Excel")
st.markdown("Faça o upload de um arquivo PDF para extrair informações de notas fiscais/faturas e exportar para uma planilha.")

uploaded_file = st.file_uploader("Escolha um arquivo PDF", type="pdf")

if uploaded_file is not None:
    # 1. Ler o arquivo e preparar para extração
    try:
        pdf_bytes = uploaded_file.read()
        st.info("PDF carregado com sucesso. Iniciando extração...")
        
        # 2. Executar a extração
        data = extract_data(pdf_bytes)
        
        if data:
            # 3. Criar o DataFrame do Pandas
            df = pd.DataFrame([data])
            
            st.success("Extração concluída!")
            
            # 4. Exibir os resultados
            st.subheader("✅ Dados Extraídos:")
            st.dataframe(df)
            
            # 5. Gerar o download do Excel
            csv = df.to_csv(index=False).encode('utf-8')
            st.download_button(
                label="📥 Baixar Resultado em Excel (.xlsx)",
                data=df.to_excel(io.BytesIO(), index=False),
                file_name=f"{uploaded_file.name.replace('.pdf', '')}_dados.xlsx",
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )

    except Exception as e:
        st.error(f"Houve um erro geral ao processar o arquivo: {e}")

else:
    st.info("Por favor, faça o upload de um arquivo PDF para começar a extração.")
