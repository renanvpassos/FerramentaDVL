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

    # Expressões Regulares Flexíveis
    # 1. Busca Invoice após "Number / Date"
    invoice_regex = re.compile(
        r"Number\s*/\s*Date\s*[:\s]*(\d+)", re.IGNORECASE
    )

    # 2. Localiza blocos de itens iniciados por um código de 6 dígitos seguido do PartNumber de 9 dígitos
    # Captura até o próximo item de 6 dígitos ou fim da página
    item_block_regex = re.compile(
        r"(?:^\s*|\n\s*)(\d{6})\s+(\d{9})\s+([\s\S]*?)(?=(?:\n\s*\d{6}\s+\d{9})|\Z)",
        re.MULTILINE,
    )

    # 3. Extrai os 3 últimos números do bloco do item (Qty, Peso, Preço)
    # Suporta formatos: 100 | 1,000.00 | 1.000,00 | 12,5
    number_regex = re.compile(r"[-+]?\d+(?:[\.,]\d+)*")

    with pdfplumber.open(pdf_file) as pdf:
        for page_num, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if not text:
                continue

            # Busca se há uma nova Invoice nesta página
            inv_match = invoice_regex.search(text)
            if inv_match:
                current_invoice = inv_match.group(1)

            # Busca todos os blocos de itens
            for match in item_block_regex.finditer(text):
                part_number = match.group(2)
                block_content = match.group(3)

                # Busca todos os números no conteúdo restante do item
                numbers = number_regex.findall(block_content)

                # Se encontrarmos pelo menos os 3 valores finais (Qty, Net, Total)
                if len(numbers) >= 3:
                    qty = numbers[-3]
                    peso_total = numbers[-2]
                    preco_total = numbers[-1]

                    records.append({
                        "Invoice": (
                            current_invoice if current_invoice else "N/A"
                        ),
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
        st.dataframe(df, use_container_width=True)

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
        st.error(
            "Nenhum dado foi encontrado com a nova estrutura. Veja abaixo o texto bruto extraído para diagnosticar."
        )

        # Ferramenta de diagnóstico inline se falhar novamente
        with pdfplumber.open(uploaded_file) as pdf:
            sample_text = (
                pdf.pages[0].extract_text()
                if len(pdf.pages) > 0
                else "PDF sem texto"
            )
            st.subheader("Texto bruto extraído da 1ª página:")
            st.code(sample_text)
