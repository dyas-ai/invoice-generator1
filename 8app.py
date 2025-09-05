import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import io
from datetime import datetime

# ===== Flexible Preprocess Excel =====
def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    df_raw = pd.read_excel(uploaded_file, header=None)

    header_row_idx = None
    stacked_header_idx = None

    # Detect header row (row containing "Style")
    for i in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[i].astype(str)
        if row.str.contains("Style", case=False, na=False).any():
            header_row_idx = i
            stacked_header_idx = i - 1
            break

    if header_row_idx is None:
        raise ValueError("❌ Could not detect header row with 'Style' column!")

    # Combine stacked headers
    if stacked_header_idx >= 0:
        headers = (
            df_raw.iloc[stacked_header_idx].astype(str).fillna('')
            + ' '
            + df_raw.iloc[header_row_idx].astype(str).fillna('')
        )
    else:
        headers = df_raw.iloc[header_row_idx].astype(str).fillna('')
    headers = headers.str.strip()

    # Stronger mapping (cover typos & variations)
    col_map = {
        "STYLE NO": ["Style", "Style No", "Item Style"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc", "Desc"],
        "COMPOSITION": ["Composition", "Fabric Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD", "Fob $"],
        "QTY": ["Total Qty", "Quantity", "Qty"],
        "AMOUNT": ["Total Value", "Amount", "Value"],
    }

    df_columns = {}
    for target_col, variants in col_map.items():
        for var in variants:
            matched_cols = [c for c in headers if var.lower() in str(c).lower()]
            if matched_cols:
                df_columns[target_col] = matched_cols[0]
                break
        if target_col not in df_columns:
            st.warning(f"⚠️ Could not find column for **{target_col}** in Excel")
            df_columns[target_col] = None

    # Select data rows
    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Rename columns
    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    df = df.rename(columns=rename_dict)

    # Drop junk rows
    if "STYLE NO" in df.columns:
        df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
        df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None", "NONE"])]
        df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]

    # Convert numerics
    if "QTY" in df.columns:
        df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)
    else:
        df["QTY"] = 0

    if "UNIT PRICE" in df.columns:
        df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0.0)
    else:
        df["UNIT PRICE"] = 0.0

    if "ITEM DESCRIPTION" not in df.columns:
        df["ITEM DESCRIPTION"] = ""

    # Aggregate
    grouped = (
        df.groupby(["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"], dropna=False)
        .agg({"QTY": "sum"})
        .reset_index()
    )

    # Compute amount
    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

    # Static columns
    grouped["FABRIC TYPE"] = "Knitted"  # default fallback
    grouped["HS CODE"] = "61112000"
    grouped["COUNTRY OF ORIGIN"] = "India"

    grouped = grouped[
        [
            "STYLE NO",
            "ITEM DESCRIPTION",
            "FABRIC TYPE",
            "HS CODE",
            "COMPOSITION",
            "COUNTRY OF ORIGIN",
            "QTY",
            "UNIT PRICE",
            "AMOUNT",
        ]
    ]
    return grouped

# ===== PDF Generator =====
def generate_proforma_invoice(df, form_data):
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    # Header
    elements.append(Paragraph("PROFORMA INVOICE", styles["Title"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"PI No & Date: {form_data['pi_number']}", styles["Normal"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph("Supplier: SAR APPARELS INDIA PVT.LTD.", styles["Normal"]))
    elements.append(Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", styles["Normal"]))
    elements.append(Paragraph("Phone: 9874173373", styles["Normal"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Buyer: {form_data['buyer_name']}", styles["Normal"]))
    elements.append(Paragraph(f"Consignee: {form_data['consignee_name']}", styles["Normal"]))
    elements.append(Paragraph(f"Address: {form_data['consignee_address']}", styles["Normal"]))
    elements.append(Paragraph(f"Tel/Fax: {form_data['consignee_tel']}", styles["Normal"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Brand Name: {form_data['brand_name']}", styles["Normal"]))
    elements.append(Paragraph(f"Payment Term: {form_data['payment_term']}", styles["Normal"]))
    elements.append(Paragraph("Port of Loading: Mumbai", styles["Normal"]))
    elements.append(Paragraph("Loading Country: India", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Table
    headers = df.columns.tolist()
    table_data = [headers]
    for _, row in df.iterrows():
        table_data.append(
            [
                row["STYLE NO"],
                row["ITEM DESCRIPTION"],
                row["FABRIC TYPE"],
                row["HS CODE"],
                row["COMPOSITION"],
                row["COUNTRY OF ORIGIN"],
                int(row["QTY"]),
                f"{row['UNIT PRICE']:.2f}",
                f"{row['AMOUNT']:.2f}",
            ]
        )
    table = Table(table_data, repeatRows=1)
    table.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.grey),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.whitesmoke),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 8),
            ]
        )
    )
    elements.append(table)
    elements.append(Spacer(1, 12))

    # Totals
    total_qty = df["QTY"].sum()
    total_amount = df["AMOUNT"].sum()
    elements.append(Paragraph(f"Total Quantity: {total_qty}", styles["Normal"]))
    elements.append(Paragraph(f"TOTAL USD {total_amount:,.2f}", styles["Normal"]))
    elements.append(Spacer(1, 12))

    # Footer
    elements.append(Paragraph("Bank: Kotak Mahindra Bank Ltd", styles["Normal"]))
    elements.append(Paragraph("SWIFT: KKBKINBBCPC", styles["Normal"]))
    elements.append(Spacer(1, 24))
    elements.append(Paragraph("Signed by: __________________", styles["Normal"]))
    elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", styles["Normal"]))

    doc.build(elements)
    buffer.seek(0)
    return buffer

# ===== Streamlit App =====
st.set_page_config(page_title="Proforma Invoice Generator", layout="centered")
st.title("📄 Proforma Invoice Generator")

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    try:
        df = preprocess_excel_flexible_auto(uploaded_file)
        st.write("### Preview of Processed Data")
        st.dataframe(df)

        st.write("### ✍️ Enter Invoice Details")
        with st.form("invoice_form"):
            pi_number = st.text_input("PI No. & Date", f"SAR/LG/XXXX Dt. {datetime.today().strftime('%d/%m/%Y')}")
            consignee_name = st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)")
            consignee_address = st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE")
            consignee_tel = st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66")
            buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
            brand_name = st.text_input("Brand Name", "Juniors")
            payment_term = st.text_input("Payment Term", "T/T")
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            form_data = {
                "pi_number": pi_number,
                "consignee_name": consignee_name,
                "consignee_address": consignee_address,
                "consignee_tel": consignee_tel,
                "buyer_name": buyer_name,
                "brand_name": brand_name,
                "payment_term": payment_term,
            }
            pdf_buffer = generate_proforma_invoice(df, form_data)
            st.success("✅ PDF Generated Successfully!")
            st.download_button(
                label="⬇️ Download Proforma Invoice",
                data=pdf_buffer,
                file_name="Proforma_Invoice.pdf",
                mime="application/pdf",
            )
    except Exception as e:
        st.error(f"❌ Error: {e}")
