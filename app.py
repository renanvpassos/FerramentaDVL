import streamlit as st
import pdfplumber
import pandas as pd
import re
import io
import pytesseract
from pdf2image import convert_from_bytes
from PIL import Image

def extract_text_ocr(pdf_bytes):
    """Converte páginas PDF para imagem e roda OCR com Tesseract."""
    # Transforma PDF em lista de imagens PIL (300 DPI para melhor precisão)
    images = convert_from_bytes(pdf_bytes, dpi=300)
    pages_ocr_text = []

    for idx, img in enumerate(images):
        # Roda OCR especificando Português e Inglês
        text = pytesseract.image_to_string(img, lang='por+eng')
        pages_ocr_text.append({"pagina": idx + 1, "texto": text})

    return pages_ocr_text

def parse_ocr_records(pages_ocr_text):
    """Processa o texto gerado pelo OCR aplicando Regex."""
    records = []
    current_invoice = None

    # Regex flexibilizados para tolerar pequenas falhas do OCR
    invoice_pattern = re.compile(r'Number\s*[/|\\]?\s*Date\s+([A-Za-z0-9\-_]+)', re.IGNORECASE)
    
    # Busca padrão de itens: PartNumber de 9 dígitos + valores
    item_pattern = re.compile(
        r'\b\d{6}\b\s+(\d{9})\s+(.*?)\s+([\d\.,]+)\s+([\d\.,]+)\s+([\d\.,]+)$'
    )

    for page_data in pages_ocr_text:
        lines = page_data["texto"].split('\n')
        for line in lines:
            line_clean = line.strip()
            if not line_clean:
                continue

            # Busca Invoice
            inv_match = invoice_pattern.search(line_clean)
            if inv_match:
                current_invoice = inv_match.group(1)
                continue

            # Busca Item
            item_match = item_pattern.search(line_clean)
            if item_match and current_invoice:
                records.append({
                    "Invoice": current_invoice,
                    "PartNumber": item_match.group(1),
                    "Quantidade": item_match.group(3),
                    "PesoTotal": item_match.group(4),
                    "PrecoTotal": item_match.group(5)
                })

    return pd.DataFrame(records)

# Interface Streamlit
st.set_page_config(page_title="Extrator PDF via OCR", layout="wide")
st.title("👁️ Extrator de PDF com OCR (Reconhecimento Óptico)")

uploaded_file = st.file_uploader("Upload do arquivo PDF", type=["pdf"])

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()

    tab_ocr, tab_result = st.tabs(["🔍 Visualização do Texto Reconhecido (OCR)", "📊 Planilha Final Extraída"])

    with st.spinner("Executando OCR nas páginas do PDF (processando imagens e caracteres)..."):
        pages_ocr = extract_text_ocr(pdf_bytes)
        df_final = parse_ocr_records(pages_ocr)

    with tab_ocr:
        st.subheader("Texto Lido via OCR por Página")
        st.caption("Confira aqui se o motor OCR conseguiu identificar corretamente os números e palavras.")
        for page in pages_ocr:
            with st.expander(f"Página {page['pagina']}", expanded=(page['pagina'] == 1)):
                st.code(page["texto"], language="text")

    with tab_result:
        if not df_final.empty:
            st.success(f"Extração concluída com sucesso! {len(df_final)} registros encontrados.")
            st.dataframe(df_final, use_container_width=True)

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df_final.to_excel(writer, index=False, sheet_name='Invoices_OCR')
            buffer.seek(0)

            st.download_button(
                label="📥 Baixar Planilha Excel",
                data=buffer,
                file_name="relatorio_invoices_ocr.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("Nenhum item no padrão foi identificado. Verifique o texto extraído na aba 'Visualização' para ajustar os critérios.")
