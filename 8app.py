import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import io

# ===== Preprocess Excel Data =====
def preprocess_excel(df):
    """
    Cleans and aggregates Excel data based on the mapping:
    STYLE NO → Style
    ITEM DESCRIPTION → Descreption
    FABRIC TYPE → (blank)
    HS CODE → (blank)
    COMPOSITION → Composition
    COUNTRY OF ORIGIN → "India"
    QTY → sum of Total Qty
    UNIT PRICE → Fob$
    AMOUNT → recomputed = Qty × Unit Price
    """

    # Rename columns to standard names
    df = df.rename(columns={
        "Style": "STYLE NO",
        "Descreption": "ITEM DESCRIPTION",   # handle Excel typo
        "Composition": "COMPOSITION",
        "Fob$": "UNIT PRICE",
        "Total Qty": "QTY",
        "Total Value": "AMOUNT"
    })

    # Convert numeric fields safely
    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0.0)

    # Group by style + description + composition + unit price
    grouped = (
        df.groupby(["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"], dropna=False)
        .agg({"QTY": "sum"})
        .reset_index()
    )

    # Compute AMOUNT
    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

    # Add static columns
    grouped["FABRIC TYPE"] = ""
    grouped["HS CODE"] = ""
    grouped["COUNTRY OF ORIGIN"] = "India"

    # Reorder columns for PDF
    grouped = grouped[
        ["STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
         "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "AMOUNT"]
    ]

    return grouped


# ===== PDF Generator =====
def generate_proforma_invoice(df):
    df = preprocess_excel(df)  # 🔥 apply mapping + aggregation
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    # --- Header ---
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

    # --- Table ---
    headers = df.columns.tolist()
    table_data = [headers]

    for _, row in df.iterrows():
        table_data.append([
            row["STYLE NO"],
            row["ITEM DESCRIPTION"],
            row["FABRIC TYPE"],
            row["HS CODE"],
            row["COMPOSITION"],
            row["COUNTRY OF ORIGIN"],
            int(row["QTY"]),
            f"{row['UNIT PRICE']:.2f}",
            f"{row['AMOUNT']:.2f}"
        ])

    table = Table(table_data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # --- Totals ---
    total_qty = df["QTY"].sum()
    total_amount = df["AMOUNT"].sum()

    elements.append(Paragraph(f"Total Quantity: {total_qty}", styles['Normal']))
    elements.append(Paragraph(f"TOTAL USD {total_amount:,.2f}", styles['Normal']))
    elements.append(Spacer(1, 12))

    # --- Footer ---
    elements.append(Paragraph("Bank: Kotak Mahindra Bank Ltd", styles['Normal']))
    elements.append(Paragraph("SWIFT: KKBKINBBCPC", styles['Normal']))
    elements.append(Spacer(1, 24))
    elements.append(Paragraph("Signed by: __________________", styles['Normal']))
    elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", styles['Normal']))

    doc.build(elements)
    buffer.seek(0)
    return buffer


# ===== Streamlit App =====
st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("📄 Proforma Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.write("### Preview of Uploaded Data")
    st.dataframe(df)

    if st.button("Generate PDF"):
        pdf_buffer = generate_proforma_invoice(df)
        st.success("✅ PDF Generated Successfully!")

        st.download_button(
            label="⬇️ Download Proforma Invoice",
            data=pdf_buffer,
            file_name="Proforma_Invoice.pdf",
            mime="application/pdf"
        )
