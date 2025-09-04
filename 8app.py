import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import io

# ===== PDF Generator =====
def generate_proforma_invoice(df, output_buffer):
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(output_buffer, pagesize=A4)
    elements = []

    # ----- Header -----
    elements.append(Paragraph("PROFORMA INVOICE", styles['Title']))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Supplier: SAR APPARELS INDIA PVT.LTD.", styles['Normal']))
    elements.append(Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", styles['Normal']))
    elements.append(Paragraph("Phone: 9874173373", styles['Normal']))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Buyer: LANDMARK GROUP", styles['Normal']))
    elements.append(Paragraph("Consignee: RNA Resources Group Ltd - Landmark (Babyshop), Dubai, UAE", styles['Normal']))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Brand Name: Juniors", styles['Normal']))
    elements.append(Paragraph("Payment Term: T/T", styles['Normal']))
    elements.append(Paragraph("Port of Loading: Mumbai", styles['Normal']))
    elements.append(Paragraph("Loading Country: India", styles['Normal']))
    elements.append(Spacer(1, 12))

    # ----- Table -----
    table_data = [["STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
                   "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE FOB", "AMOUNT"]]

    total_qty = df["Total Qty"].sum()
    total_value = df["Total Value"].sum()

    for _, row in df.iterrows():
        table_data.append([
            row["Style"],
            row["Description"],
            "KNITTED",       # fixed
            "61112000",      # fixed
            row["Composition"],
            "India",         # fixed
            int(row["Total Qty"]),
            f"{row['USD Fob$']:.2f}",
            f"{row['Total Value']:.2f}"
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.grey),
        ("TEXTCOLOR", (0,0), (-1,0), colors.whitesmoke),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 8),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # ----- Totals -----
    elements.append(Paragraph(f"Total Quantity: {int(total_qty)}", styles['Normal']))
    elements.append(Paragraph(f"TOTAL USD {total_value:,.2f}", styles['Normal']))
    elements.append(Spacer(1, 12))

    # ----- Footer -----
    elements.append(Paragraph("Bank: Kotak Mahindra Bank Ltd", styles['Normal']))
    elements.append(Paragraph("SWIFT: KKBKINBBCPC", styles['Normal']))
    elements.append(Spacer(1, 24))
    elements.append(Paragraph("Signed by: __________________", styles['Normal']))
    elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", styles['Normal']))

    # Build PDF
    doc.build(elements)

# ===== Streamlit App =====
st.title("📄 Proforma Invoice Generator")

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
        output_buffer = io.BytesIO()
        generate_proforma_invoice(df, output_buffer)
        st.success("✅ PDF generated successfully!")

        st.download_button(
            label="📥 Download Invoice PDF",
            data=output_buffer.getvalue(),
            file_name="Proforma_Invoice.pdf",
            mime="application/pdf"
        )
