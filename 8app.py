import streamlit as st
from fpdf import FPDF
import io

# ---------------- PDF Class ----------------
class InvoicePDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "PROFORMA INVOICE", new_x="LMARGIN", new_y="NEXT", align="C")
        self.ln(5)

    def add_static_info(self):
        self.set_font("Arial", "", 10)

        # Supplier / Exporter
        self.multi_cell(0, 6, "Supplier Name: SAR APPARELS INDIA PVT.LTD.")
        self.multi_cell(0, 6, "ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.")
        self.multi_cell(0, 6, "PHONE : 9874173373")
        self.multi_cell(0, 6, "FAX : N.A.")
        self.ln(2)

        # Buyer / Consignee
        self.multi_cell(0, 6, "Buyer Name: LANDMARK GROUP")
        self.multi_cell(0, 6, "Brand Name: Juniors")
        self.multi_cell(0, 6, "Consignee: RNA Resources Group Ltd- Landmark (Babyshop)")
        self.multi_cell(0, 6, "P O Box 25030, Dubai, UAE")
        self.multi_cell(0, 6, "Tel: 00971 4 8095500")
        self.multi_cell(0, 6, "Fax: 00971 4 8095555/66")
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
        self.multi_cell(0, 6, "Bank Details (Including Swift/IBAN)")
        self.multi_cell(0, 6, "BENEFICIARY: SAR APPARELS INDIA PVT.LTD")
        self.multi_cell(0, 6, "ACCOUNT NO: 2112819952")
        self.multi_cell(0, 6, "BANK NAME: KOTAK MAHINDRA BANK LTD")
        self.multi_cell(0, 6, "BANK ADDRESS: 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001")
        self.multi_cell(0, 6, "SWIFT CODE: KKBKINBBCPC")
        self.multi_cell(0, 6, "BANK CODE: 0323")
        self.ln(5)

    def add_table(self):
        self.set_font("Arial", "B", 9)
        headers = [
            "STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE",
            "KNITTED / WOVEN", "H.S NO (8digit)", "COMPOSITION OF MATERIAL",
            "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "FOB", "AMOUNT"
        ]
        col_widths = [20, 35, 20, 25, 25, 35, 25, 15, 20, 15, 25]

        # Table header
        for i, h in enumerate(headers):
            self.multi_cell(col_widths[i], 10, h, border=1, align="C", max_line_height=self.font_size, new_x="RIGHT", new_y="TOP")
        self.ln()

        # Table rows (dummy static data)
        self.set_font("Arial", "", 9)
        rows = [
            ["SAV001S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,107", "6.00", "", "24,642.00"],
            ["SAV002S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,593", "6.00", "", "27,558.00"],
            ["SAV003S25", "S/L Bodysuit 7pk", "KNITTED", "61112000", "100% COTTON", "India", "4,593", "6.00", "", "27,558.00"],
        ]

        for row in rows:
            for i, item in enumerate(row):
                self.multi_cell(col_widths[i], 8, str(item), border=1, align="C", max_line_height=self.font_size, new_x="RIGHT", new_y="TOP")
            self.ln()

        # Total
        self.set_font("Arial", "B", 10)
        self.cell(0, 10, "USD 79,758.00", new_x="LMARGIN", new_y="NEXT", align="R")
        self.multi_cell(0, 8, "TOTAL US DOLLAR SEVENTY-NINE THOUSAND SEVEN HUNDRED FIFTY-EIGHT DOLLARS")
        self.ln(5)
        self.multi_cell(0, 8, "Signed by ………………(…Affix Stamp here) for RNA Resources Group Ltd-Landmark (Babyshop)")
        self.ln(10)
        self.multi_cell(0, 8, "Terms & Conditions (If Any)")

# ---------------- PDF Generator ----------------
def generate_invoice():
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()
    pdf.add_static_info()
    pdf.add_table()

    pdf_bytes = pdf.output(dest="S")
    return io.BytesIO(pdf_bytes)

# ---------------- Streamlit App ----------------
def main():
    st.title("📄 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File (not yet used, static demo)", type=["xlsx"])
    if uploaded_file:
        st.write("✅ File uploaded successfully!")

        pdf_buffer = generate_invoice()

        st.success("✅ Static invoice generated!")
        st.download_button(
            "⬇️ Download Invoice PDF",
            data=pdf_buffer,
            file_name="invoice.pdf",
            mime="application/pdf"
        )

if __name__ == "__main__":
    main()
