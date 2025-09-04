import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# ===== PDF Generator (fpdf2) =====
class InvoicePDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 14)
        self.cell(0, 10, "PROFORMA INVOICE", ln=True, align="C")
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")


def generate_proforma_invoice(df):
    pdf = InvoicePDF("P", "mm", "A4")
    pdf.add_page()

    # ===== Supplier / Buyer Info =====
    pdf.set_font("Helvetica", size=10)
    pdf.cell(0, 6, "Supplier: SAR APPARELS INDIA PVT.LTD.", ln=True)
    pdf.cell(0, 6, "Address: 6, Picaso Bithi, Kolkata - 700017", ln=True)
    pdf.cell(0, 6, "Phone: 9874173373", ln=True)
    pdf.ln(5)

    pdf.cell(0, 6, "Buyer: LANDMARK GROUP", ln=True)
    pdf.multi_cell(0, 6, "Consignee: RNA Resources Group Ltd - Landmark (Babyshop), Dubai, UAE")
    pdf.ln(5)

    pdf.cell(0, 6, "Brand Name: Juniors", ln=True)
    pdf.cell(0, 6, "Payment Term: T/T", ln=True)
    pdf.cell(0, 6, "Port of Loading: Mumbai", ln=True)
    pdf.cell(0, 6, "Loading Country: India", ln=True)
    pdf.ln(5)

    # ===== Table Header =====
    col_widths = [28, 35, 25, 20, 32, 28, 15, 18, 25]  # widths for each column
    headers = ["STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
               "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "AMOUNT"]

    pdf.set_fill_color(60, 60, 60)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 8)
    for i, header in enumerate(headers):
        pdf.cell(col_widths[i], 8, header, border=1, align="C", fill=True)
    pdf.ln()

    # ===== Table Rows =====
    pdf.set_font("Helvetica", size=8)
    pdf.set_text_color(0, 0, 0)

    total_qty = df["Total Qty"].sum()
    total_value = df["Total Value"].sum()

    for _, row in df.iterrows():
        row_data = [
            str(row["Style"]),
            str(row["Description"]),
            "KNITTED",       # fixed
            "61112000",      # fixed
            str(row["Composition"]),
            "India",         # fixed
            str(int(row["Total Qty"])),
            f"{row['USD Fob$']:.2f}",
            f"{row['Total Value']:.2f}"
        ]
        for i, item in enumerate(row_data):
            pdf.cell(col_widths[i], 8, item, border=1, align="C")
        pdf.ln()

    pdf.ln(5)

    # ===== Totals =====
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(0, 6, f"Total Quantity: {int(total_qty)}", ln=True)
    pdf.cell(0, 6, f"TOTAL USD {total_value:,.2f}", ln=True)
    pdf.ln(10)

    # ===== Footer / Bank Info =====
    pdf.set_font("Helvetica", size=9)
    pdf.cell(0, 6, "Bank: Kotak Mahindra Bank Ltd", ln=True)
    pdf.cell(0, 6, "SWIFT: KKBKINBBCPC", ln=True)
    pdf.ln(10)

    pdf.cell(0, 6, "Signed by: __________________", ln=True)
    pdf.multi_cell(0, 6, "For RNA Resources Group Ltd - Landmark (Babyshop)")

    return pdf.output(dest="S").encode("latin1")


# ===== STREAMLIT APP =====
st.title("📄 Proforma Invoice Generator (FPDF2)")

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    # ---- Read Excel with multi-row headers ----
    df = pd.read_excel(uploaded_file, header=[0,1])
    df.columns = [' '.join(col).strip() for col in df.columns.values]  # flatten

    # ---- Fix column names ----
    if "Descreption " in df.columns:
        df = df.rename(columns={"Descreption ": "Description"})
    if "Material Composition" in df.columns:
        df = df.rename(columns={"Material Composition": "Composition"})

    # ---- Keep required columns ----
    df = df[["Style", "Description", "Composition", "USD Fob$", "Total Qty", "Total Value"]]

    # ---- Group by Style ----
    df = df.groupby(["Style", "Description", "Composition", "USD Fob$"], as_index=False).agg({
        "Total Qty": "sum",
        "Total Value": "sum"
    })

    st.write("✅ Processed Invoice Data:")
    st.dataframe(df)

    # ---- PDF Generation ----
    if st.button("Generate PDF"):
        pdf_bytes = generate_proforma_invoice(df)
        st.success("✅ PDF generated successfully!")

        st.download_button(
            label="📥 Download Invoice PDF",
            data=pdf_bytes,
            file_name="Proforma_Invoice.pdf",
            mime="application/pdf"
        )
