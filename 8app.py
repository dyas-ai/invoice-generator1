import streamlit as st
import pandas as pd
from fpdf import FPDF
import io

# -------------------------------
# Helper: PDF Class
# -------------------------------
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Proforma Invoice", ln=1, align="C")
        self.ln(2)

    def add_static_info(self, header):
        self.set_font("Arial", "", 9)

        # Supplier + PI info (two-column layout)
        self.cell(95, 6, f"Supplier Name: {header['Supplier Name']}", border=1)
        self.cell(95, 6, f"No. & date of PI: {header['PI Number']} Dt. {header['PI Date']}", border=1, ln=1)

        self.cell(95, 6, f"Address: {header['Supplier Address']}", border=1)
        self.cell(95, 6, f"Order Reference: {header['Order Reference']}", border=1, ln=1)

        self.cell(95, 6, f"Phone: {header['Phone']}", border=1)
        self.cell(95, 6, f"Buyer Name: {header['Buyer Name']}", border=1, ln=1)

        self.cell(95, 6, f"Fax: {header['Fax']}", border=1)
        self.cell(95, 6, f"Brand Name: {header['Brand Name']}", border=1, ln=1)

        # Consignee & Payment info
        self.multi_cell(95, 6, f"Consignee:\n{header['Consignee']}", border=1)
        self.set_xy(105, self.get_y() - 12)
        self.multi_cell(95, 6, f"Payment Term: {header['Payment Term']}", border=1)

        # Shipment info
        self.cell(95, 6, f"Loading Country: {header['Loading Country']}", border=1)
        self.cell(95, 6, f"Port of Loading: {header['Port of Loading']}", border=1, ln=1)

        self.cell(95, 6, f"Agreed Shipment Date: {header['Agreed Shipment Date']}", border=1)
        self.cell(95, 6, f"Description of Goods: {header['Description of Goods']}", border=1, ln=1)

        self.ln(5)

    def add_table(self, df):
        self.set_font("Arial", "B", 8)
        col_widths = [20, 40, 25, 20, 20, 20, 20, 20, 20, 25]

        headers = [
            "Style No.",
            "Item Description",
            "Fabric Type",
            "H.S No",
            "Composition",
            "Country of Origin",
            "Qty",
            "Unit Price",
            "FOB",
            "Amount",
        ]

        for i, h in enumerate(headers):
            self.cell(col_widths[i], 6, h, border=1, align="C")
        self.ln()

        self.set_font("Arial", "", 8)
        for _, row in df.iterrows():
            self.cell(col_widths[0], 6, str(row.get("Style", "")), border=1)
            self.cell(col_widths[1], 6, str(row.get("Description", ""))[:25], border=1)
            self.cell(col_widths[2], 6, "Knitted", border=1)  # Static for now
            self.cell(col_widths[3], 6, "61091000", border=1)  # Example HS code
            self.cell(col_widths[4], 6, str(row.get("Composition", "")), border=1)
            self.cell(col_widths[5], 6, "India", border=1)  # Static
            self.cell(col_widths[6], 6, str(row.get("Total Qty", "")), border=1, align="R")
            self.cell(col_widths[7], 6, str(row.get("USD Fob$", "")), border=1, align="R")
            self.cell(col_widths[8], 6, str(row.get("USD Fob$", "")), border=1, align="R")
            self.cell(col_widths[9], 6, str(row.get("Total Value", "")), border=1, align="R")
            self.ln()

# -------------------------------
# Extract static fields from Excel
# -------------------------------
def extract_static_info(df):
    return {
        "Supplier Name": "SAR APPARELS INDIA PVT.LTD.",
        "Supplier Address": "6, Picaso Bithi, KOLKATA - 700017",
        "Phone": "9874173373",
        "Fax": "N.A.",
        "PI Number": "SAR/LG/0148",
        "PI Date": "14-10-2024",
        "Order Reference": "CPO/47062/25",
        "Buyer Name": "LANDMARK GROUP",
        "Brand Name": "Juniors",
        "Consignee": "RNA Resources Group Ltd - Landmark (Babyshop), P O Box 25030, Dubai, UAE",
        "Payment Term": "T/T",
        "Loading Country": "India",
        "Port of Loading": "Mumbai",
        "Agreed Shipment Date": "07-02-2025",
        "Description of Goods": "Value Packs",
    }

# -------------------------------
# Main Streamlit App
# -------------------------------
def main():
    st.title("📑 Proforma Invoice Generator")

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])
    if uploaded_file:
        df = pd.read_excel(uploaded_file, header=[9])  # Table starts at row 10
        df = df.rename(columns=lambda x: str(x).strip())

        # Extract required columns safely
        table_df = df[["Style", "Description", "Composition", "USD Fob$", "Total Qty", "Total Value"]].copy()

        # Group by Style
        table_df = table_df.groupby(["Style", "Description", "Composition", "USD Fob$"], as_index=False).sum()

        header_info = extract_static_info(df)

        # Generate PDF
        pdf = PDF()
        pdf.add_page()
        pdf.add_static_info(header_info)
        pdf.add_table(table_df)

        pdf_bytes = pdf.output(dest="S").encode("latin1")
        st.download_button("⬇️ Download Invoice PDF", data=pdf_bytes, file_name="invoice.pdf", mime="application/pdf")

if __name__ == "__main__":
    main()
