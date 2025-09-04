import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

class InvoicePDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "Proforma Invoice", border=1, new_x="LMARGIN", new_y="NEXT", align="C")

    def add_static_info(self, info):
        self.set_font("Helvetica", "", 10)
        # Left box
        self.multi_cell(95, 6,
            f"Supplier Name\nSAR APPARELS INDIA PVT.LTD.\n"
            f"ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.\n"
            f"PHONE : 9874173373\nFAX : N.A.",
            border=1)
        # Right box
        self.set_xy(105, 20)
        self.multi_cell(95, 6,
            f"No. & date of PI : {info.get('PI No', 'SAR/LG/0148 Dt. 14-10-2024')}\n"
            f"Landmark order Reference: {info.get('Order Ref', 'CPO/47062/25')}\n"
            f"Buyer Name: LANDMARK GROUP\n"
            f"Brand Name: Juniors",
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

def generate_invoice(info):
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()
    pdf.add_static_info(info)
    pdf.add_table()

    # ✅ Correct Streamlit-compatible PDF return
    pdf_bytes = pdf.output(dest="S").encode("latin-1")
    return io.BytesIO(pdf_bytes)

# ------------------- Streamlit UI -------------------
def main():
    st.title("📄 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx", "xls"])

    if st.button("Generate Invoice"):
        # Right now we ignore Excel, just make static invoice
        info = {
            "PI No": "SAR/LG/0148 Dt. 14-10-2024",
            "Order Ref": "CPO/47062/25"
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
