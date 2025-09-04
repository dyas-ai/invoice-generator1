import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# --------------------
# PDF Generator
# --------------------
def generate_invoice(df, pi_number="PI-001", po_number="PO-001", date="2025-09-04"):
    pdf = FPDF("P", "mm", "A4")
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 14)

    # ---- Title ----
    pdf.cell(0, 10, "PROFORMA INVOICE", ln=True, align="C")

    pdf.set_font("Helvetica", "", 10)
    pdf.ln(5)
    pdf.cell(0, 6, f"PI No: {pi_number}", ln=True)
    pdf.cell(0, 6, f"PO No: {po_number}", ln=True)
    pdf.cell(0, 6, f"Date: {date}", ln=True)

    pdf.ln(8)

    # ---- Parties ----
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(95, 6, "Supplier:", border=1)
    pdf.cell(95, 6, "Buyer:", border=1, ln=True)

    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(95, 6, "Your Company Name\nAddress Line 1\nCity, Country", border=1)
    x = pdf.get_x()
    y = pdf.get_y() - 18
    pdf.set_xy(x+95, y)
    pdf.multi_cell(95, 6, "Buyer Name\nBuyer Address Line 1\nCity, Country", border=1)

    pdf.ln(5)

    # ---- Bank Details ----
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, "Bank Details:", ln=True)
    pdf.set_font("Helvetica", "", 10)
    pdf.multi_cell(0, 6, "Bank Name: Example Bank\nAccount No: 123456789\nSWIFT: EXAMPLExx\nBranch: City Branch")

    pdf.ln(8)

    # ---- Table Header ----
    pdf.set_font("Helvetica", "B", 8)
    col_widths = [20, 30, 22, 22, 22, 35, 28, 15, 20, 25]  
    headers = [
        "STYLE NO.",
        "ITEM DESCRIPTION",
        "FABRIC TYPE",
        "KNITTED/WOVEN",
        "H.S NO (8digit)",
        "COMPOSITION OF MATERIAL",
        "COUNTRY OF ORIGIN",
        "QTY",
        "UNIT PRICE FOB",
        "AMOUNT"
    ]

    for i, header in enumerate(headers):
        pdf.multi_cell(col_widths[i], 10, header, border=1, align="C", max_line_height=4)
        x = pdf.get_x() + col_widths[i]
        y = pdf.get_y() - 10
        pdf.set_xy(x, y)
    pdf.ln()

    # ---- Table Data from Excel ----
    pdf.set_font("Helvetica", "", 8)
    for _, row in df.iterrows():
        row_data = [
            str(row.get("Style", "")),
            str(row.get("Description", "")),
            str(row.get("Fabric Type", "")),
            str(row.get("Knitted/Woven", "")),
            str(row.get("HS Code", "")),
            str(row.get("Composition", "")),
            str(row.get("Country of Origin", "")),
            str(row.get("Qty", "")),
            str(row.get("Unit Price", "")),
            str(row.get("Amount", ""))
        ]
        for i, val in enumerate(row_data):
            pdf.cell(col_widths[i], 8, val, border=1, align="C")
        pdf.ln()

    # ---- Footer / Signature ----
    pdf.ln(10)
    pdf.cell(0, 6, "For Your Company Name", ln=True, align="R")
    pdf.ln(15)
    pdf.cell(0, 6, "Authorised Signatory", ln=True, align="R")

    # ---- Output ----
    pdf_bytes = pdf.output(dest="S")
    return io.BytesIO(pdf_bytes)


# --------------------
# Streamlit UI
# --------------------
st.title("📄 Proforma Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, header=[0,1])
        df.columns = [' '.join([str(c) for c in col]).strip() for col in df.columns.values]
    except Exception:
        df = pd.read_excel(uploaded_file)

    st.write("✅ Excel loaded:", df.head())

    if st.button("Generate Invoice PDF"):
        pdf_file = generate_invoice(df)
        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_file,
            file_name="proforma_invoice.pdf",
            mime="application/pdf"
        )
