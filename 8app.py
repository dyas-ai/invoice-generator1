import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
from datetime import datetime
from num2words import num2words

# ===== Preprocess Excel =====
def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    df_raw = pd.read_excel(uploaded_file, header=None)

    header_row_idx = None
    stacked_header_idx = None
    for i in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[i].astype(str)
        if row.str.contains("Style", case=False, na=False).any():
            header_row_idx = i
            stacked_header_idx = i - 1
            break

    if header_row_idx is None:
        raise ValueError("Could not detect header row with 'Style' column!")

    if stacked_header_idx >= 0:
        headers = (
            df_raw.iloc[stacked_header_idx].astype(str).fillna('')
            + ' '
            + df_raw.iloc[header_row_idx].astype(str).fillna('')
        )
    else:
        headers = df_raw.iloc[header_row_idx].astype(str).fillna('')

    headers = headers.str.strip()

    col_map = {
        "STYLE NO": ["Style", "Style No", "Item Style"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc"],
        "COMPOSITION": ["Composition", "Fabric Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD", "Fob $"],
        "QTY": ["Total Qty", "Quantity", "Qty"],
        "AMOUNT": ["Total Value", "Amount", "Value"],
        "FABRIC TYPE": ["Texture", "Fabric", "Knitted", "Woven"]
    }

    df_columns = {}
    for target_col, variants in col_map.items():
        for var in variants:
            matched_cols = [c for c in headers if var.lower() in str(c).lower()]
            if matched_cols:
                df_columns[target_col] = matched_cols[0]
                break
        if target_col not in df_columns:
            df_columns[target_col] = None

    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    df = df.rename(columns=rename_dict)

    if "STYLE NO" in df.columns:
        df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
        df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None", "NONE"])]
        df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]

    df["QTY"] = pd.to_numeric(df.get("QTY", 0), errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df.get("UNIT PRICE", 0), errors="coerce").fillna(0.0)

    df = df[~((df["QTY"] == 0) & (df["UNIT PRICE"] == 0) & (df["STYLE NO"].str.strip() == ""))]

    grouped = (
        df.groupby(["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE", "FABRIC TYPE"], dropna=False)
        .agg({"QTY": "sum"})
        .reset_index()
    )

    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

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
def generate_proforma_invoice(df, details):
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    bold = ParagraphStyle("Bold", parent=normal, fontName="Helvetica-Bold")
    underline = ParagraphStyle("Underline", parent=normal, fontName="Helvetica-Bold", underline=True)

    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=30, bottomMargin=30)
    elements = []

    # Header
    elements.append(Paragraph("PROFORMA INVOICE", styles["Title"]))
    elements.append(Spacer(1, 6))
    elements.append(Paragraph(f"Proforma Invoice No.: {details['pi_no']}", normal))
    elements.append(Paragraph(f"Date: {details['pi_date']}", normal))
    elements.append(Paragraph(f"Order Reference: {details['order_ref']}", normal))
    elements.append(Paragraph(f"Buyer Name: {details['buyer_name']}", normal))
    elements.append(Paragraph(f"Brand Name: {details['brand_name']}", normal))
    elements.append(Spacer(1, 12))

    # Supplier & Consignee
    elements.append(Paragraph(f"Supplier: {details['supplier_name']}", normal))
    elements.append(Paragraph(f"Address: {details['supplier_address']}", normal))
    elements.append(Paragraph(f"Phone: {details['supplier_phone']}", normal))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Consignee: {details['consignee_name']}", normal))
    elements.append(Paragraph(f"Address: {details['consignee_address']}", normal))
    elements.append(Paragraph(f"Phone/Fax: {details['consignee_phone']}", normal))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"Payment Term: {details['payment_term']}", normal))
    elements.append(Spacer(1, 12))

    # Shipment Details
    elements.append(Paragraph(f"Loading Country: {details['loading_country']}", normal))
    elements.append(Paragraph(f"Port of Loading: {details['port_of_loading']}", normal))
    elements.append(Paragraph(f"Shipment Date: {details['shipment_date']}", normal))
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
                ("ALIGN", (6, 1), (-1, -1), "RIGHT"),
                ("ALIGN", (0, 0), (-1, 0), "CENTER"),
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
    elements.append(Paragraph(f"TOTAL QTY: {total_qty}", bold))
    elements.append(Paragraph(f"TOTAL USD {total_amount:,.2f}", bold))

    # Amount in words - bold + underline
    amount_words = num2words(total_amount, to="currency", lang="en").upper()
    elements.append(Paragraph(f"Amount in words: {amount_words}", underline))
    elements.append(Spacer(1, 12))

    # Bank Details
    elements.append(Paragraph("Bank Details (Including Swift/IBAN)", bold))
    elements.append(Paragraph(f"Bank: {details['bank_name']}", normal))
    elements.append(Paragraph(f"Branch: {details['bank_branch']}", normal))
    elements.append(Paragraph(f"SWIFT: {details['bank_swift']}", normal))
    elements.append(Paragraph(f"IBAN: {details['bank_iban']}", normal))
    elements.append(Paragraph(f"Account Number: {details['bank_account']}", normal))
    elements.append(Spacer(1, 12))

    # Footer
    elements.append(Spacer(1, 24))
    elements.append(Paragraph("Signed by: __________________ (Affix Stamp here)", normal))
    elements.append(Paragraph("For RNA Resources Group Ltd - Landmark (Babyshop)", normal))

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

        # Invoice Details Form
        with st.form("invoice_form"):
            st.subheader("Invoice Details")
            pi_no = st.text_input("Proforma Invoice No.", "SAR/LG/0148")
            pi_date = datetime.today().strftime("%d-%b-%Y")
            order_ref = st.text_input("Order Reference", "")
            buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
            brand_name = st.text_input("Brand Name", "Juniors")
            supplier_name = st.text_input("Supplier Name", "SAR APPARELS INDIA PVT.LTD.")
            supplier_address = st.text_input("Supplier Address", "6, Picaso Bithi, Kolkata - 700017")
            supplier_phone = st.text_input("Supplier Phone", "9874173373")
            consignee_name = st.text_input("Consignee Name", "RNA Resources Group Ltd - Landmark (Babyshop)")
            consignee_address = st.text_area("Consignee Address", "Dubai, UAE")
            consignee_phone = st.text_input("Consignee Phone/Fax", "")
            payment_term = st.text_input("Payment Term", "T/T")
            loading_country = st.text_input("Loading Country", "India")
            port_of_loading = st.text_input("Port of Loading", "Mumbai")
            shipment_date = st.text_input("Shipment Date", "")
            st.subheader("Bank Details")
            bank_name = st.text_input("Bank Name", "Kotak Mahindra Bank Ltd")
            bank_branch = st.text_input("Branch", "")
            bank_swift = st.text_input("SWIFT", "KKBKINBBCPC")
            bank_iban = st.text_input("IBAN", "")
            bank_account = st.text_input("Account Number", "")
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            details = {
                "pi_no": pi_no,
                "pi_date": pi_date,
                "order_ref": order_ref,
                "buyer_name": buyer_name,
                "brand_name": brand_name,
                "supplier_name": supplier_name,
                "supplier_address": supplier_address,
                "supplier_phone": supplier_phone,
                "consignee_name": consignee_name,
                "consignee_address": consignee_address,
                "consignee_phone": consignee_phone,
                "payment_term": payment_term,
                "loading_country": loading_country,
                "port_of_loading": port_of_loading,
                "shipment_date": shipment_date,
                "bank_name": bank_name,
                "bank_branch": bank_branch,
                "bank_swift": bank_swift,
                "bank_iban": bank_iban,
                "bank_account": bank_account,
            }
            pdf_buffer = generate_proforma_invoice(df, details)
            st.success("✅ PDF Generated Successfully!")
            st.download_button(
                label="⬇️ Download Proforma Invoice",
                data=pdf_buffer,
                file_name=f"Proforma_Invoice_{pi_no}.pdf",
                mime="application/pdf",
            )
    except Exception as e:
        st.error(f"❌ Error: {e}")
