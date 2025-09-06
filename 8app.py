import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
from num2words import num2words

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
        "STYLE NO": ["Style", "Style No", "Item Style"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc"],
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
            df_columns[target_col] = None  

    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    df = df.rename(columns=rename_dict)

    if "STYLE NO" in df.columns:
        df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
        df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None", "NONE"])]
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
def generate_proforma_invoice(df, form_inputs):
    buffer = io.BytesIO()
    styles = getSampleStyleSheet()
    normal = styles["Normal"]
    bold = ParagraphStyle("Bold", parent=normal, fontName="Helvetica-Bold")
    underline = ParagraphStyle("Underline", parent=normal, underline=True)

    doc = SimpleDocTemplate(buffer, pagesize=A4)
    elements = []

    # Header
    elements.append(Paragraph("PROFORMA INVOICE", styles["Title"]))
    elements.append(Spacer(1, 12))
    elements.append(Paragraph(f"PI No. & Date: {form_inputs['pi_number_date']}", normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph("Supplier: SAR APPARELS INDIA PVT.LTD.", normal))
    elements.append(Paragraph("Address: 6, Picaso Bithi, Kolkata - 700017", normal))
    elements.append(Paragraph("Phone: 9874173373", normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph(f"Consignee: {form_inputs['consignee_name']}", normal))
    elements.append(Paragraph(f"Address: {form_inputs['consignee_address']}", normal))
    elements.append(Paragraph(f"Tel/Fax: {form_inputs['consignee_tel']}", normal))
    elements.append(Spacer(1, 12))

    elements.append(Paragraph(f"Buyer: {form_inputs['buyer_name']}", normal))
    elements.append(Paragraph(f"Brand Name: {form_inputs['brand_name']}", normal))
    elements.append(Paragraph(f"Payment Term: {form_inputs['payment_term']}", normal))
    elements.append(Paragraph(f"Port of Loading: {form_inputs['port_loading']}", normal))
    elements.append(Paragraph(f"Loading Country: {form_inputs['loading_country']}", normal))
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
    total_qty = int(df["QTY"].sum())
    total_amount = float(df["AMOUNT"].sum())  # ✅ FIX applied here
    elements.append(Paragraph(f"TOTAL QTY: {total_qty}", bold))
    elements.append(Paragraph(f"TOTAL USD {total_amount:,.2f}", bold))

    amount_words = num2words(total_amount, to="currency", lang="en").upper()
    elements.append(Paragraph(f"Amount in words: {amount_words}", underline))
    elements.append(Spacer(1, 12))

    # Bank details
    elements.append(Paragraph(f"Bank: {form_inputs['bank_name']}", normal))
    elements.append(Paragraph(f"SWIFT: {form_inputs['swift']}", normal))
    elements.append(Paragraph(f"IBAN: {form_inputs['iban']}", normal))
    elements.append(Paragraph(f"Account Number: {form_inputs['account_number']}", normal))
    elements.append(Spacer(1, 24))

    elements.append(Paragraph("Signed by: __________________", normal))
    elements.append(Paragraph(f"For {form_inputs['consignee_name']}", normal))

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

        st.subheader("✍️ Enter Invoice Details")
        form_inputs = {
            "pi_number_date": st.text_input("PI No. & Date", "SAR/LG/XXXX Dt. 04/09/2025"),
            "consignee_name": st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)"),
            "consignee_address": st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE"),
            "consignee_tel": st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66"),
            "buyer_name": st.text_input("Buyer Name", "LANDMARK GROUP"),
            "brand_name": st.text_input("Brand Name", "Juniors"),
            "payment_term": st.text_input("Payment Term", "T/T"),
            "port_loading": st.text_input("Port of Loading", "Mumbai"),
            "loading_country": st.text_input("Loading Country", "India"),
            "bank_name": st.text_input("Bank Name", "Kotak Mahindra Bank Ltd"),
            "swift": st.text_input("SWIFT", "KKBKINBBCPC"),
            "iban": st.text_input("IBAN", "123456789"),
            "account_number": st.text_input("Account Number", "987654321"),
        }

        if st.button("Generate PDF"):
            pdf_buffer = generate_proforma_invoice(df, form_inputs)
            st.success("✅ PDF Generated Successfully!")
            st.download_button(
                label="⬇️ Download Proforma Invoice",
                data=pdf_buffer,
                file_name="Proforma_Invoice.pdf",
                mime="application/pdf",
            )
    except Exception as e:
        st.error(f"❌ Error: {e}")
