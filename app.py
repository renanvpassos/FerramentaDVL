import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

def inspect_pdf_text(pdf_file):
    """Extrai todas as linhas de texto cruas marcadas por página para visualização."""
    raw_pages = []
    with pdfplumber.open(pdf_file) as pdf:
        for idx, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            lines = [line.strip() for line in text.split('\n') if line.strip()]
            raw_pages.append({"pagina": idx + 1, "linhas": lines})
    return raw_pages

def parse_pdf(pdf_file):
    records = []
    current_invoice = None

    # Regex Patterns
    invoice_pattern = re.compile(r'Number\s*/\s*Date\s+([A-Za-z0-9\-_]+)', re.IGNORECASE)
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

                inv_match = invoice_pattern.search(line_clean)
                if inv_match:
                    current_invoice = inv_match.group(1)
                    continue

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
st.set_page_config(page_title="Leitor e Extrator de PDF", layout="wide")
st.title("📄 Inspeção e Extração de PDF para Excel")

uploaded_file = st.file_uploader("Selecione o arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    # Criação de abas para separar a pré-visualização da tabela final
    tab_inspect, tab_result = st.tabs(["🔍 Inspeção das Linhas Lidas (Pré-processamento)", "📊 Tabela Final Extraída"])

    with tab_inspect:
        st.subheader("Texto Cru Extraído por Página")
        st.caption("Verifique aqui como o leitor lê as linhas do PDF para entender se o layout está correto.")
        
        pages_data = inspect_pdf_text(uploaded_file)
        
        for page in pages_data:
            with st.expander(f"Página {page['pagina']} ({len(page['linhas'])} linhas detectadas)", expanded=(page['pagina'] == 1)):
                # Exibe em formato de tabela interativa para fácil inspeção
                df_lines = pd.DataFrame(page['linhas'], columns=["Linha de Texto Capturada"])
                st.dataframe(df_lines, use_container_width=True)

    with tab_result:
        uploaded_file.seek(0)  # Reseta o ponteiro do arquivo
        with st.spinner("Processando extração..."):
            df = parse_pdf(uploaded_file)

        if not df.empty:
            st.success(f"{len(df)} itens identificados e estruturados com sucesso!")
            st.dataframe(df, use_container_width=True)

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
            st.error("Nenhum item foi capturado. Compare os padrões das linhas na aba 'Inspeção' para ajustar os filtros.")
