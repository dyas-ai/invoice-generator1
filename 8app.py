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

    def add_static_info(self):
        self.set_font("Arial", "", 10)

        # Supplier / Exporter
        self.multi_cell(0, 6, "Supplier Name: SAR APPARELS INDIA PVT.LTD.")
        self.multi_cell(0, 6, "ADDRESS : 6, Picaso Bithi, KOLKATA - 700017")
        self.multi_cell(0, 6, "PHONE : 9874173373")
        self.multi_cell(0, 6, "FAX : N.A.")
        self.ln(2)

        # Buyer / Consignee
        self.multi_cell(0, 6, "Buyer Name: LANDMARK GROUP")
        self.multi_cell(0, 6, "Brand Name: Juniors")
        self.multi_cell(0, 6, "Consignee: RNA Resources Group Ltd- Landmark (Babyshop)")
        self.multi_cell(0, 6, "Address: P O Box 25030, Dubai, UAE")
        self.multi_cell(0, 6, "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66")
        self.ln(2)

        # PI / Order details
        self.multi_cell(0, 6, "No. & date of PI: SAR/LG/0148 Dt. 14-10-2024")
        self.multi_cell(0, 6, "Landmark Order Reference: CPO/47062/25")
        self.multi_cell(0, 6, "Payment Term: T/T")
        self.multi_cell(0, 6, "Port of Loading: Mumbai")
        self.multi_cell(0, 6, "Loading Country: India")
        self.multi_cell(0, 6, "Agreed Shipment Date: 07-02-2025")
        self.ln(2)

        # Bank details
        self.multi_cell(0, 6, "Bank Details:")
        self.multi_cell(0, 6, "BENEFICIARY: SAR APPARELS INDIA PVT.LTD")
        self.multi_cell(0, 6, "ACCOUNT NO: 2112819952")
        self.multi_cell(0, 6, "BANK NAME: KOTAK MAHINDRA BANK LTD")
        self.multi_cell(0, 6, "BANK ADDRESS: 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001")
        self.multi_cell(0, 6, "SWIFT CODE: KKBKINBBCPC")
        self.multi_cell(0, 6, "BANK CODE: 0323")
        self.ln(5)

    def add_table(self, items):
        headers = [
            "STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE",
            "H.S NO (8 digit)", "COMPOSITION OF MATERIAL",
            "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE FOB", "AMOUNT"
        ]

        col_widths = [22, 30, 22, 22, 30, 25, 15, 22, 25]

        # Table header
        self.set_font("Arial", "B", 9)
        for i, header in enumerate(headers):
            self.cell(col_widths[i], 10, header, border=1, align="C")
        self.ln()

        # Table rows
        self.set_font("Arial", "", 9)
        for row in items:
            for i, value in enumerate(row):
                self.cell(col_widths[i], 8, str(value), border=1, align="C")
            self.ln()

    def add_totals(self):
        self.ln(5)
        self.set_font("Arial", "B", 10)
        self.cell(0, 8, "TOTAL: USD 79,758.00", ln=1, align="R")
        self.set_font("Arial", "I", 9)
        self.multi_cell(0, 6, "TOTAL US DOLLAR SEVENTY-NINE THOUSAND SEVEN HUNDRED FIFTY-EIGHT DOLLARS", align="R")
        self.ln(5)
        self.set_font("Arial", "", 9)
        self.cell(0, 8, "Signed by ……………… (Affix Stamp here) for RNA Resources Group Ltd-Landmark (Babyshop)", ln=1)

# -------------------------------
# Generate Invoice
# -------------------------------
def generate_invoice():
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()

    pdf.add_static_info()

    # Dummy table data (from your example)
    items = [
        ["SAV001S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4107", "6.00", "24642.00"],
        ["SAV002S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4593", "6.00", "27558.00"],
        ["SAV003S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4593", "6.00", "27558.00"],
    ]
    pdf.add_table(items)
    pdf.add_totals()

    pdf_bytes = pdf.output(dest="S")
    return io.BytesIO(pdf_bytes)

# -------------------------------
# Streamlit UI
# -------------------------------
def main():
    st.title("📑 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file)
        st.write("✅ Excel file uploaded successfully!")
        st.dataframe(df.head())

        pdf_buffer = generate_invoice()

        st.success("✅ Invoice generated!")
        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_buffer,
            file_name="invoice.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
