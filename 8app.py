import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

class InvoicePDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "Proforma Invoice", border=1, ln=1, align="C")

    def add_supplier_block(self):
        self.set_font("Helvetica", "", 10)
        self.multi_cell(95, 6,
            "Supplier Name\nSAR APPARELS INDIA PVT.LTD.\n"
            "ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.\n"
            "PHONE : 9874173373\n"
            "FAX : N.A.",
            border=1)

    def add_invoice_details_block(self):
        self.set_xy(105, 20)
        self.multi_cell(95, 6,
            "No. & date of PI : SAR/LG/0148 Dt. 14-10-2024\n"
            "Landmark order Reference: CPO/47062/25\n"
            "Buyer Name: LANDMARK GROUP\n"
            "Brand Name: Juniors",
            border=1)

    def add_consignee_payment(self):
        self.set_xy(10, 60)
        self.multi_cell(95, 6,
            "Consignee:-\nRNA Resources Group Ltd - Landmark (Babyshop),\n"
            "P O Box 25030, Dubai, UAE,\n"
            "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66",
            border=1)

        self.set_xy(105, 60)
        self.multi_cell(95, 6,
            "Payment Term: T/T\n\n"
            "Bank Details (Including Swift/IBAN)\n"
            "BENIFICIARY :- SAR APPARELS INDIA PVT.LTD\n"
            "ACCOUNT NO :- 2112819952\n"
            "BANK'S NAME :- KOTAK MAHINDRA BANK LTD\n"
            "BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR,\n"
            "                KOLKATA-700001\n"
            "SWIFT CODE :- KKBKINBBCPC\n"
            "BANK CODE :- 0323",
            border=1)

    def add_shipping_block(self):
        self.set_xy(10, 110)
        self.multi_cell(95, 6,
            "Loading Country: India\n"
            "Port of loading: Mumbai\n"
            "Agreed Shipment Date: 07-02-2025",
            border=1)

        self.set_xy(105, 110)
        self.multi_cell(95, 6,
            "L/C Advising Bank (If Payment term LC Applicable )\n\n"
            "REMARKS if ANY:-",
            border=1)

    def add_goods_description(self):
        self.set_xy(10, 140)
        self.multi_cell(190, 6,
            "Description of goods: Value Packs",
            border=1)

    def add_table(self):
        headers = [
            "STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE",
            "KNITTED/WOVEN", "H.S NO (8digit)", "COMPOSITION",
            "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE FOB", "AMOUNT"
        ]
        col_widths = [20, 40, 25, 25, 25, 30, 30, 20, 25, 25]

        self.set_font("Helvetica", "B", 8)
        for h, w in zip(headers, col_widths):
            self.cell(w, 8, h, border=1, align="C")
        self.ln()

        self.set_font("Helvetica", "", 8)
        for i in range(5):  # placeholder rows
            for w in col_widths:
                self.cell(w, 8, "", border=1)
            self.ln()

def generate_static_invoice():
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()
    pdf.add_supplier_block()
    pdf.add_invoice_details_block()
    pdf.add_consignee_payment()
    pdf.add_shipping_block()
    pdf.add_goods_description()
    pdf.add_table()

    # Output to memory instead of file
    pdf_bytes = pdf.output(dest="S").encode("latin1")
    return pdf_bytes

# ---------------- Streamlit UI ----------------
def main():
    st.title("📄 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])
    
    if st.button("Generate Invoice"):
        pdf_bytes = generate_static_invoice()
        st.success("✅ Invoice generated!")

        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_bytes,
            file_name="invoice.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
