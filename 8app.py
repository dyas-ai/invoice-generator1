import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import io

# ===== Flexible Preprocess Excel =====
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
        "STYLE NO": ["Style", "Style No"],
        "ITEM DESCRIPTION": ["Descreption", "Description"],
        "COMPOSITION": ["Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD"],
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
            df_columns[target_col] = None

    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    df = df.rename(columns=rename_dict)

    if "STYLE NO" in df.columns:
        df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
        df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None"])]
        df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]

    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0.0)

    df = df[~((df["QTY"] == 0) & (df["UNIT PRICE"] == 0) & (df["STYLE NO"].str.strip() == ""))]

    grouped = (
        df.groupby(["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"], dropna=False)
        .agg({"QTY": "sum"})
        .reset_index()
    )

    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]
    grouped["FABRIC TYPE"] = "Knitted"
    grouped["HS CODE"] = "61112000"
    grouped["COUNTRY OF ORIGIN"] = "India"

    grouped = grouped[
        [
            "STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
            "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "AMOUNT"
        ]
    ]

    return grouped

# ===== PDF Generator with Boxed Layout =====
def generate_proforma_invoice(df, form_data):
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=20, bottomMargin=20, leftMargin=20, rightMargin=20)
    elements = []

    elements.append(Paragraph("<b>PROFORMA INVOICE</b>", styles["Title"]))
    elements.append(Spacer(1, 6))

    # Supplier & PI Info
    supplier_info = [
        Paragraph("<b>Supplier Name:</b>", styles["Normal"]),
        Paragraph("<b>SAR APPARELS INDIA PVT.LTD.</b>", styles["Normal"]),
        Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", styles["Normal"]),
        Paragraph("Phone: 9874173373", styles["Normal"]),
        Paragraph("Fax: N.A.", styles["Normal"]),
    ]
    pi_info = [
        Paragraph(f"<b>No. & date of PI:</b> {form_data['pi_number']}", styles["Normal"]),
        Paragraph(f"<b>Landmark order Reference:</b> {form_data['order_ref']}", styles["Normal"]),
        Paragraph(f"<b>Buyer Name:</b> {form_data['buyer_name']}", styles["Normal"]),
        Paragraph(f"<b>Brand Name:</b> {form_data['brand_name']}", styles["Normal"]),
    ]
    block1 = Table([[supplier_info, pi_info]], colWidths=[270, 270])
    block1.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.75, colors.black), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    elements.append(block1)

    # Consignee & Bank
    consignee_info = [
        Paragraph("<b>Consignee:</b>", styles["Normal"]),
        Paragraph(form_data['consignee_name'], styles["Normal"]),
        Paragraph(form_data['consignee_address'], styles["Normal"]),
        Paragraph(form_data['consignee_tel'], styles["Normal"]),
    ]
    bank_info = [
        Paragraph(f"<b>Payment Term:</b> {form_data['payment_term']}", styles["Normal"]),
        Paragraph("<b>Bank Details (Including Swift/IBAN)</b>", styles["Normal"]),
        Paragraph(f"<b>Beneficiary:</b> {form_data['bank_beneficiary']}", styles["Normal"]),
        Paragraph(f"<b>Account No:</b> {form_data['bank_account']}", styles["Normal"]),
        Paragraph(f"<b>BANK'S NAME:</b> {form_data['bank_name']}", styles["Normal"]),
        Paragraph(f"<b>BANK ADDRESS:</b> {form_data['bank_address']}", styles["Normal"]),
        Paragraph(f"<b>SWIFT CODE:</b> {form_data['bank_swift']}", styles["Normal"]),
        Paragraph(f"<b>BANK CODE:</b> {form_data['bank_code']}", styles["Normal"]),
    ]
    block2 = Table([[consignee_info, bank_info]], colWidths=[270, 270])
    block2.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.75, colors.black), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    elements.append(block2)

    # Shipment & Remarks
    shipment_info = [
        Paragraph(f"<b>Loading Country:</b> {form_data['loading_country']}", styles["Normal"]),
        Paragraph(f"<b>Port of Loading:</b> {form_data['port_loading']}", styles["Normal"]),
        Paragraph(f"<b>Agreed Shipment Date:</b> {form_data['shipment_date']}", styles["Normal"]),
    ]
    remarks_info = [
        Paragraph(f"<b>L/C Advising Bank:</b> (If applicable)", styles["Normal"]),
        Paragraph(f"<b>Remarks:</b> {form_data['remarks']}", styles["Normal"]),
    ]
    block3 = Table([[shipment_info, remarks_info]], colWidths=[270, 270])
    block3.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.75, colors.black), ("VALIGN", (0,0), (-1,-1), "TOP")]))
    elements.append(block3)

    # Goods
    block4 = Table([[Paragraph(f"<b>Description of goods:</b> {form_data['goods_desc']}", styles["Normal"])]],
                   colWidths=[540])
    block4.setStyle(TableStyle([("GRID", (0,0), (-1,-1), 0.75, colors.black)]))
    elements.append(block4)

    elements.append(Spacer(1, 12))

    # Line Items
    headers = df.columns.tolist()
    table_data = [headers]
    for _, row in df.iterrows():
        table_data.append([
            row["STYLE NO"], row["ITEM DESCRIPTION"], row["FABRIC TYPE"],
            row["HS CODE"], row["COMPOSITION"], row["COUNTRY OF ORIGIN"],
            int(row["QTY"]), f"{row['UNIT PRICE']:.2f}", f"{row['AMOUNT']:.2f}"
        ])
    table = Table(table_data, repeatRows=1, colWidths=[65, 110, 70, 70, 80, 80, 40, 50, 60])
    table.setStyle(TableStyle([
        ("GRID", (0,0), (-1,-1), 0.5, colors.black),
        ("BACKGROUND", (0,0), (-1,0), colors.lightgrey),
        ("ALIGN", (0,0), (-1,-1), "CENTER"),
        ("VALIGN", (0,0), (-1,-1), "MIDDLE"),
        ("FONTNAME", (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE", (0,0), (-1,-1), 7),
    ]))
    elements.append(table)

    total_qty = df["QTY"].sum()
    total_amount = df["AMOUNT"].sum()
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"<b>Total Quantity:</b> {total_qty}", styles["Normal"]))
    elements.append(Paragraph(f"<b>TOTAL USD {total_amount:,.2f}</b>", styles["Normal"]))

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

        st.write("### 📝 Enter Invoice Details")
        form_data = {
            "pi_number": st.text_input("PI No. & Date", "SAR/LG/XXXX Dt. 05/09/2025"),
            "order_ref": st.text_input("Order Reference", "CPO/47062/25"),
            "buyer_name": st.text_input("Buyer Name", "LANDMARK GROUP"),
            "brand_name": st.text_input("Brand Name", "Juniors"),
            "consignee_name": st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)"),
            "consignee_address": st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE"),
            "consignee_tel": st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66"),
            "payment_term": st.text_input("Payment Term", "T/T"),
            "bank_beneficiary": st.text_input("Bank Beneficiary", "SAR APPARELS INDIA PVT.LTD"),
            "bank_account": st.text_input("Bank Account", "2112819952"),
            "bank_name": st.text_input("Bank Name", "KOTAK MAHINDRA BANK LTD"),
            "bank_address": st.text_area("Bank Address", "2 Brabourne Road, Govind Bhavan, Ground Floor, Kolkata-700001"),
            "bank_swift": st.text_input("Bank Swift Code", "KKBKINBBCPC"),
            "bank_code": st.text_input("Bank Code", "0323"),
            "loading_country": st.text_input("Loading Country", "India"),
            "port_loading": st.text_input("Port of Loading", "Mumbai"),
            "shipment_date": st.text_input("Agreed Shipment Date", "07/02/2025"),
            "remarks": st.text_area("Remarks", ""),
            "goods_desc": st.text_input("Description of Goods", "Value Packs"),
        }

        if st.button("Generate PDF"):
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
