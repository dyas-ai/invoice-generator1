import streamlit as st
import pandas as pd
from fpdf import FPDF
import io
from datetime import datetime

# ---------------- PDF Generator ----------------
class PDF(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 12)
        self.cell(0, 10, "Proforma Invoice", ln=True, align="C")

    def footer(self):
        self.set_y(-15)
        self.set_font("Helvetica", "I", 8)
        self.cell(0, 10, f"Page {self.page_no()}", align="C")

def generate_invoice(df, pi_number=None, po_number=None):
    pdf = PDF()
    pdf.add_page()
    pdf.set_font("Helvetica", size=10)

    # ---- Header Info ----
    pdf.cell(0, 10, f"PI Number: {pi_number or 'N/A'}", ln=True)
    pdf.cell(0, 10, f"PO Number: {po_number or 'N/A'}", ln=True)
    pdf.cell(0, 10, f"Date: {datetime.today().strftime('%d-%m-%Y')}", ln=True)
    pdf.ln(5)

    # ---- Table Header ----
    col_names = list(df.columns)
    col_widths = [30, 40, 30, 25, 25, 30]  # adjust widths
    for i, col in enumerate(col_names):
        pdf.cell(col_widths[i % len(col_widths)], 8, col, border=1, align="C")
    pdf.ln()

    # ---- Table Rows ----
    for _, row in df.iterrows():
        for i, col in enumerate(col_names):
            pdf.cell(col_widths[i % len(col_widths)], 8, str(row[col]), border=1)
        pdf.ln()

    # ---- Totals ----
    if "Total Qty" in df.columns:
        pdf.ln(5)
        pdf.set_font("Helvetica", "B", 10)
        pdf.cell(0, 10, f"Grand Total Qty: {df['Total Qty'].sum()}", ln=True)
    if "Total Value" in df.columns:
        pdf.cell(0, 10, f"Grand Total Value: {df['Total Value'].sum():,.2f}", ln=True)

    # ---- Save PDF ----
    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)
    return buffer

# ---------------- Streamlit App ----------------
st.title("📄 Excel → Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    # ---- Read multi-row headers ----
    df = pd.read_excel(uploaded_file, header=[0, 1])

    # ---- Flatten headers ----
    df.columns = [
        " ".join([str(c).strip() for c in col if str(c).strip() != ""])
        for col in df.columns.values
    ]

    # ---- Normalize headers ----
    df.columns = (
        df.columns.str.strip()
        .str.lower()
        .str.replace("$", "", regex=False)
        .str.replace("  ", " ")
    )

    # ---- Map variations → standard ----
    col_map = {
        "style": "Style",
        "style no": "Style",
        "styleno": "Style",
        "description": "Description",
        "descreption": "Description",
        "composition": "Composition",
        "material composition": "Composition",
        "usd fob": "USD Fob$",
        "fob usd": "USD Fob$",
        "usd fob$": "USD Fob$",
        "total qty": "Total Qty",
        "total quantity": "Total Qty",
        "quantity total": "Total Qty",
        "total value": "Total Value",
        "value total": "Total Value",
        "totalvalue": "Total Value",
    }
    df = df.rename(columns=lambda c: col_map.get(c, c))

    # ---- Required Columns ----
    required_cols = ["Style", "Description", "Composition", "USD Fob$", "Total Qty", "Total Value"]
    if not set(required_cols[:-1]).issubset(df.columns):
        st.error(f"❌ Missing required columns in Excel. Found: {list(df.columns)}")
        st.stop()

    # ---- Auto-calc Total Value if missing ----
    if "Total Value" not in df.columns:
        df["Total Value"] = df["Total Qty"] * df["USD Fob$"]
    else:
        df["Total Value"] = df["Total Value"].fillna(df["Total Qty"] * df["USD Fob$"])

    # ---- Group by Style, Description, Composition, Price ----
    df = df.groupby(
        ["Style", "Description", "Composition", "USD Fob$"],
        as_index=False
    ).agg({
        "Total Qty": "sum",
        "Total Value": "sum"
    })

    # Reorder columns
    df = df[required_cols]

    st.success("✅ Excel processed & grouped successfully!")
    st.dataframe(df)

    # ---- Input fields ----
    pi_number = st.text_input("Enter PI Number")
    po_number = st.text_input("Enter PO Number")

    # ---- Generate PDF ----
    if st.button("Generate Invoice PDF"):
        buffer = generate_invoice(df, pi_number, po_number)
        st.download_button(
            label="⬇️ Download Invoice",
            data=buffer,
            file_name=f"Invoice_{pi_number or 'NA'}.pdf",
            mime="application/pdf",
        )
