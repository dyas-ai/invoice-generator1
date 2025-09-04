import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime

# ---------------- PDF GENERATOR ----------------
def generate_invoice(df, pi_number="PI-0001", po_number="PO-0001", buyer="Sample Buyer", seller="Sample Seller"):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Arial", "B", 14)

    # ---- Title ----
    pdf.cell(200, 10, "PROFORMA INVOICE", ln=True, align="C")

    pdf.set_font("Arial", size=10)
    pdf.cell(100, 8, f"Proforma Invoice No: {pi_number}", ln=0)
    pdf.cell(100, 8, f"Date: {datetime.today().strftime('%d-%m-%Y')}", ln=1)

    pdf.cell(100, 8, f"Purchase Order No: {po_number}", ln=0)
    pdf.cell(100, 8, f"Buyer: {buyer}", ln=1)
    pdf.cell(100, 8, f"Seller: {seller}", ln=1)
    pdf.ln(5)

    # ---- Table Header ----
    pdf.set_font("Arial", "B", 9)
    headers = ["Style", "Description", "Composition", "USD Fob$", "Total Qty", "Total Value"]
    col_widths = [25, 55, 40, 25, 25, 25]

    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 8, header, border=1, align="C")
    pdf.ln()

    # ---- Table Rows ----
    pdf.set_font("Arial", size=9)
    for _, row in df.iterrows():
        pdf.cell(col_widths[0], 8, str(row["Style"]), border=1)
        pdf.cell(col_widths[1], 8, str(row["Description"]), border=1)
        pdf.cell(col_widths[2], 8, str(row["Composition"]), border=1)
        pdf.cell(col_widths[3], 8, str(row["USD Fob$"]), border=1, align="R")
        pdf.cell(col_widths[4], 8, str(row["Total Qty"]), border=1, align="R")
        pdf.cell(col_widths[5], 8, str(row["Total Value"]), border=1, align="R")
        pdf.ln()

    # ---- Totals ----
    pdf.set_font("Arial", "B", 10)
    pdf.cell(145, 8, "TOTAL", border=1, align="R")
    pdf.cell(25, 8, str(df["Total Qty"].sum()), border=1, align="R")
    pdf.cell(25, 8, str(df["Total Value"].sum()), border=1, align="R")
    pdf.ln(15)

    pdf.cell(0, 8, "Authorized Signatory", ln=True, align="R")

    # ---- Return Bytes ----
    pdf_output = io.BytesIO()
    pdf.output(pdf_output)
    pdf_output.seek(0)
    return pdf_output


# ---------------- STREAMLIT APP ----------------
st.title("📄 Proforma Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel Invoice", type=["xlsx"])

if uploaded_file:
    # ---- Read Excel with multi-row headers ----
    df = pd.read_excel(uploaded_file, header=[0, 1])
    df.columns = [' '.join([str(c) for c in col]).strip() for col in df.columns.values]

    # ---- Normalize headers ----
    rename_map = {
        "Descreption ": "Description",
        "Descreption": "Description",
        "Material Composition": "Composition",
        "USD FOB": "USD Fob$",
        "USD FOB$": "USD Fob$",
        "USD FOB $": "USD Fob$",
    }
    df = df.rename(columns={k: v for k, v in rename_map.items() if k in df.columns})

    # ---- Required Columns ----
    required_cols = ["Style", "Description", "Composition", "USD Fob$", "Total Qty", "Total Value"]
    missing = [col for col in required_cols if col not in df.columns]

    if missing:
        st.error(f"❌ Missing required columns in Excel: {missing}")
        st.stop()

    # ---- Keep only needed columns ----
    df = df[required_cols]

    # ---- Show Processed Data ----
    st.write("✅ Processed Invoice Data:")
    st.dataframe(df)

    # ---- Generate PDF ----
    if st.button("Generate PDF"):
        pdf_file = generate_invoice(df)
        st.download_button(
            label="📥 Download Invoice PDF",
            data=pdf_file,
            file_name="proforma_invoice.pdf",
            mime="application/pdf"
        )
