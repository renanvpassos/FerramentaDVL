import streamlit as st
from pdf2docx import Converter
import docx
import pandas as pd
import re
import tempfile
import os
import io

def process_pdf_via_docx(pdf_file):
    # Salva o arquivo enviado temporariamente para a conversão
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_pdf:
        tmp_pdf.write(pdf_file.read())
        tmp_pdf_path = tmp_pdf.name

    tmp_docx_path = tmp_pdf_path.replace(".pdf", ".docx")

    try:
        # Step 1: Converte PDF -> DOCX (Reconstrói tabelas automaticamente)
        cv = Converter(tmp_pdf_path)
        cv.convert(tmp_docx_path, start=0, end=None)
        cv.close()

        # Step 2: Lê o documento Word convertido
        doc = docx.Document(tmp_docx_path)
        
        preview_tables = []
        extracted_records = []
        current_invoice = None

        # Regex para identificar Invoice
        invoice_pattern = re.compile(r'Number\s*/\s*Date\s+([A-Za-z0-9\-_]+)', re.IGNORECASE)

        # Primeiro varre parágrafos fora de tabelas buscando a Invoice
        # E varre as tabelas reconstruídas
        for table_idx, table in enumerate(doc.tables):
            table_data = []
            for row in table.rows:
                row_cells = [cell.text.strip().replace('\n', ' ') for cell in row.cells]
                table_data.append(row_cells)

                row_str = " ".join(row_cells)

                # Atualiza a invoice se encontrar na tabela
                inv_match = invoice_pattern.search(row_str)
                if inv_match:
                    current_invoice = inv_match.group(1)

                # Busca o item (PartNumber de 9 dígitos)
                # Verifica se a linha possui células suficientes e o PartNumber correto
                for i, cell_text in enumerate(row_cells):
                    pn_match = re.search(r'\b\d{9}\b', cell_text)
                    if pn_match:
                        part_number = pn_match.group(0)
                        
                        # Captura valores das colunas vizinhas na mesma linha da tabela
                        # Ajuste os índices conforme a estrutura das colunas no seu documento
                        quantidade = row_cells[i+1] if i+1 < len(row_cells) else ""
                        peso_total = row_cells[i+2] if i+2 < len(row_cells) else ""
                        preco_total = row_cells[i+3] if i+3 < len(row_cells) else ""

                        extracted_records.append({
                            "Invoice": current_invoice or "N/A",
                            "PartNumber": part_number,
                            "Quantidade": quantidade,
                            "PesoTotal": peso_total,
                            "PrecoTotal": preco_total
                        })

            if table_data:
                df_table = pd.DataFrame(table_data)
                preview_tables.append((table_idx + 1, df_table))

        return preview_tables, pd.DataFrame(extracted_records)

    finally:
        # Limpeza de arquivos temporários
        if os.path.exists(tmp_pdf_path):
            os.remove(tmp_pdf_path)
        if os.path.exists(tmp_docx_path):
            os.remove(tmp_docx_path)


# Interface Streamlit
st.set_page_config(page_title="PDF via DOCX Converter", layout="wide")
st.title("📄 Extrator de PDF (via Reconstrução DOCX)")

uploaded_file = st.file_uploader("Upload do arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    tab_preview, tab_final = st.tabs(["🧩 Tabelas Reconstruídas (Visualização Word)", "📊 Planilha Final Extraída"])

    with st.spinner("Convertendo PDF para Word e detectando estrutura de tabelas..."):
        tables_preview, df_final = process_pdf_via_docx(uploaded_file)

    with tab_preview:
        st.subheader("Tabelas Detectadas Após Conversão")
        st.caption("Abaixo estão as tabelas exatamente como o Word as enxergou após a conversão do PDF.")
        if tables_preview:
            for idx, df_t in tables_preview:
                st.write(f"**Tabela #{idx}** ({len(df_t)} linhas)")
                st.dataframe(df_t, use_container_width=True)
        else:
            st.warning("Nenhuma tabela estruturada foi identificada no documento.")

    with tab_final:
        if not df_final.empty:
            st.success(f"Extração concluída! {len(df_final)} itens organizados.")
            st.dataframe(df_final, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Invoices')
            buffer.seek(0)

            st.download_button(
                label="📥 Baixar Excel Tratado",
                data=buffer,
                file_name="relatorio_invoices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Não foi possível montar a planilha final. Confira as tabelas na primeira aba para conferir em quais colunas os dados ficaram posicionados.")
