import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def parse_pdf(pdf_file):
    records = []
    current_invoice = None

    # Regex Patterns
    # Busca "Number / Date" seguido pelo número da invoice (ex: Number / Date 12345678)
    invoice_pattern = re.compile(r'Number\s*/\s*Date\s+([A-Za-z0-9\-_]+)', re.IGNORECASE)
    
    # Captura a linha do item:
    # 1. Item ID (6 dígitos) -> \b\d{6}\b
    # 2. PartNumber (9 dígitos) -> (\d{9})
    # 3. Descrição (texto) -> (.*?)
    # 4. Quantidade -> ([\d\.,]+)
    # 5. Peso Total (Net) -> ([\d\.,]+)
    # 6. Preço Total (Total) -> ([\d\.,]+)
    item_pattern = re.compile(
        r'\b\d{6}\b\s+(\d{9})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$'
    )

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')
            for line in lines:
                line_clean = line.strip()

                # Verifica se há atualização de Invoice na página/linha
                inv_match = invoice_pattern.search(line_clean)
                if inv_match:
                    current_invoice = inv_match.group(1)
                    continue

                # Busca padrão de itens
                item_match = item_pattern.search(line_clean)
                if item_match and current_invoice:
                    part_number = item_match.group(1)
                    quantidade = item_match.group(3)
                    peso_total = item_match.group(4)
                    preco_total = item_match.group(5)

                    records.append({
                        "Invoice": current_invoice,
                        "PartNumber": part_number,
                        "Quantidade": quantidade,
                        "PesoTotal": peso_total,
                        "PrecoTotal": preco_total
                    })

    return pd.DataFrame(records)

# Interface Streamlit
st.set_page_config(page_title="PDF Extractor to Excel", layout="wide")
st.title("📄 Extrator de PDF para Excel")
st.write("Faça o upload do documento PDF para extrair as Invoices e itens correspondentes.")

uploaded_file = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    with st.spinner("Processando e extraindo dados do PDF..."):
        df = parse_pdf(uploaded_file)

    if not df.empty:
        st.success(f"Extração concluída! {len(df)} itens encontrados.")
        st.dataframe(df, use_container_width=True)

        # Exportação para Excel em memória
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Invoices')
        buffer.seek(0)

        st.download_button(
            label="📥 Baixar Planilha Excel",
            data=buffer,
            file_name="relatorio_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Nenhum dado no padrão esperado foi localizado no PDF. Verifique a estrutura do documento.")
