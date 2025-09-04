import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# -------------------------------
# PDF Class
# -------------------------------
class InvoicePDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "PROFORMA INVOICE", ln=1, align="C")
        self.ln(5)

    def add_static_info(self, info):
        self.set_font("Arial", "", 10)
        for key, value in info.items():
            self.cell(50, 8, f"{key}:", border=0)
            self.cell(0, 8, str(value), border=0, ln=1)
        self.ln(5)

    def add_table(self):
        # Table headers
        headers = [
            "STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE",
            "H.S NO (8 digit)", "COMPOSITION OF MATERIAL",
            "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE FOB", "AMOUNT"
        ]
        self.set_font("Arial", "B", 9)
        for header in headers:
            self.cell(22, 10, header, border=1, align="C")
        self.ln()

        # Dummy table rows (later we will map from Excel)
        self.set_font("Arial", "", 9)
        for i in range(5):
            self.cell(22, 8, f"Style {i+1}", border=1)
            self.cell(22, 8, "T-Shirt", border=1)
            self.cell(22, 8, "Knitted", border=1)
            self.cell(22, 8, "61091000", border=1)
            self.cell(22, 8, "100% Cotton", border=1)
            self.cell(22, 8, "India", border=1)
            self.cell(22, 8, "100", border=1, align="R")
            self.cell(22, 8, "5.00", border=1, align="R")
            self.cell(22, 8, "500.00", border=1, align="R")
            self.ln()

# -------------------------------
# Generate Invoice Function
# -------------------------------
def generate_invoice(info):
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()
    pdf.add_static_info(info)
    pdf.add_table()

    # ✅ FIX: fpdf2 returns bytearray, no need to encode
    pdf_bytes = pdf.output(dest="S")
    return io.BytesIO(pdf_bytes)

# -------------------------------
# Main Streamlit App
# -------------------------------
def main():
    st.title("📑 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("✅ Excel file uploaded successfully!")
        st.dataframe(df.head())

        # Hardcoded static info for now
        info = {
            "Exporter": "Sarla Performance Fibers Ltd.",
            "Consignee": "XYZ Imports Ltd.",
            "Invoice Date": "14-10-2024",
            "PI No": "SAR/LG/0148 Dt. 14-10-2024",
            "Order Ref": "CPO/47062/25",
            "Port of Loading": "Mumbai",
            "Agreed Shipment Date": "07-02-2025",
            "Description of Goods": "Value Packs"
        }

        pdf_buffer = generate_invoice(info)

        st.success("✅ Invoice generated!")
        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_buffer,
            file_name="invoice.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
