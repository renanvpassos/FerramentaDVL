import io
import re
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Extrator de Invoice PDF", page_icon="📄", layout="wide"
)

st.title("📄 Extrator de Dados de Invoice (PDF -> Excel)")
st.write(
    "Faça o upload do seu relatório em PDF para extrair as informações e gerar a planilha Excel."
)


def extract_data_from_pdf(pdf_file):
    records = []
    current_invoice = None

    # Regex Patterns
    # Busca "Number / Date" seguido por dígitos (Invoice)
    invoice_pattern = re.compile(
        r"Number\s*/\s*Date\s*\n?\s*(\d+)", re.IGNORECASE
    )

    # Captura a linha de item:
    # 1. Item ID (6 dígitos)
    # 2. Part Number (9 dígitos)
    # 3. Descrição (texto intermediário)
    # 4. Quantidade (número inteiro ou decimal)
    # 5. Peso Total (Net - número decimal/inteiro)
    # 6. Preço Total (Total - número decimal/inteiro)
    item_pattern = re.compile(
        r"^\s*\d{6}\s+(\d{9})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)\s*$",
        re.MULTILINE,
    )

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            # Atualiza a Invoice se encontrada na página corrente
            invoice_match = invoice_pattern.search(text)
            if invoice_match:
                current_invoice = invoice_match.group(1)

            # Extrai todas as linhas de itens presentes na página
            for match in item_pattern.finditer(text):
                part_number = match.group(1)
                qty = match.group(3)
                peso_total = match.group(4)
                preco_total = match.group(5)

                records.append({
                    "Invoice": current_invoice if current_invoice else "N/A",
                    "PartNumber": part_number,
                    "Quantidade": qty,
                    "PesoTotal": peso_total,
                    "PrecoTotal": preco_total,
                })

    return pd.DataFrame(records)


# Interface de Upload
uploaded_file = st.file_uploader(
    "Selecione o arquivo PDF", type=["pdf"]
)

if uploaded_file is not None:
    with st.spinner("Processando o arquivo PDF..."):
        df = extract_data_from_pdf(uploaded_file)

    if not df.empty:
        st.success(f"Extração concluída! {len(df)} itens encontrados.")

        # Visualização da Tabela
        st.dataframe(df, use_container_width=True)

        # Download para Excel
        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="Invoices")

        st.download_button(
            label="📥 Baixar Planilha Excel",
            data=buffer.getvalue(),
            file_name="dados_invoices.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.warning(
            "Nenhum dado foi encontrado com o padrão especificado. Verifique a estrutura do PDF."
        )
