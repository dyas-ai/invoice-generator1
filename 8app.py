import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.units import inch
import io

# ===== Preprocessing Function (robust + fills missing expected cols) =====
def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    # read everything with no header (we will detect header row)
    df_raw = pd.read_excel(uploaded_file, header=None)

    # detect header row (row containing "Style" keyword)
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

    # build combined headers if there is a stacked header above
    if stacked_header_idx >= 0:
        headers = (
            df_raw.iloc[stacked_header_idx].astype(str).fillna("") + " " + df_raw.iloc[header_row_idx].astype(str).fillna("")
        )
    else:
        headers = df_raw.iloc[header_row_idx].astype(str).fillna("")
    headers = headers.str.strip().astype(str)

    # possible column name variants to match
    col_map = {
        "STYLE NO": ["Style", "Style No", "Item Style", "STYLE"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc", "DESC"],
        "COMPOSITION": ["Composition", "Fabric Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD", "Fob $", "Unit Price", "FOB"],
        "QTY": ["Total Qty", "Quantity", "Qty", "QTY"],
        "AMOUNT": ["Total Value", "Amount", "Value", "TOTAL VALUE"],
    }

    # map found header text -> standard column name
    df_columns = {}
    hdr_list = list(headers)
    for target_col, variants in col_map.items():
        found = None
        for var in variants:
            for hdr in hdr_list:
                if var.lower() in str(hdr).lower():
                    found = hdr
                    break
            if found:
                break
        df_columns[target_col] = found  # may be None

    # create dataframe of data rows beneath header row
    df = df_raw.iloc[header_row_idx + 1 :].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # rename columns that were matched
    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    if rename_dict:
        df = df.rename(columns=rename_dict)

    # Ensure required columns exist — if missing, create them with defaults
    required_cols_defaults = {
        "STYLE NO": "",
        "ITEM DESCRIPTION": "",
        "COMPOSITION": "",
        "UNIT PRICE": 0.0,
        "QTY": 0,
        "AMOUNT": 0.0,
    }
    for col, default in required_cols_defaults.items():
        if col not in df.columns:
            df[col] = default

    # Clean STYLE NO column and drop rows that are clearly not product rows
    df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
    # drop rows where style no is empty or looks like footer/notes
    df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None", "NONE"])]
    df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]

    # Convert numeric fields safely
    df["QTY"] = pd.to_numeric(df.get("QTY", 0), errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df.get("UNIT PRICE", 0.0), errors="coerce").fillna(0.0).astype(float)
    # If AMOUNT exists in raw, prefer recalculating to ensure consistency
    # (some sheets had 'Total Value' entries on separate rows — we recompute)
    df["AMOUNT"] = df["QTY"] * df["UNIT PRICE"]

    # Remove any rows that are purely zeros and have no style
    df = df[~((df["QTY"] == 0) & (df["UNIT PRICE"] == 0) & (df["STYLE NO"].str.strip() == ""))]

    # Group by style to aggregate quantities (one row per unique style)
    group_by_cols = ["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"]
    # Ensure all group_by columns exist
    for c in group_by_cols:
        if c not in df.columns:
            df[c] = "" if c != "UNIT PRICE" else 0.0

    grouped = (
        df.groupby(group_by_cols, dropna=False, as_index=False)
        .agg({"QTY": "sum"})
        .reset_index(drop=True)
    )

    # Recompute AMOUNT = QTY * UNIT PRICE (definitive)
    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

    # Extract Fabric Type (Texture) from top rows if present (safe attempt)
    fabric_type = ""
    try:
        # scan first max_rows for a cell that contains "Texture"
        search_block = df_raw.iloc[:max_rows].astype(str)
        mask = search_block.apply(lambda r: r.str.contains("Texture", case=False, na=False)).any(axis=1)
        if mask.any():
            row_idx = mask.idxmax()
            # find the column index where 'Texture' occurred
            texture_row = search_block.iloc[row_idx]
            texture_cols = [i for i, v in enumerate(texture_row) if "texture" in str(v).lower()]
            if texture_cols:
                col_idx = texture_cols[0] + 1
                # ensure col_idx in bounds
                if col_idx < df_raw.shape[1]:
                    val = df_raw.iat[row_idx, col_idx]
                    if pd.notna(val) and str(val).strip() != "":
                        fabric_type = str(val).strip()
    except Exception:
        fabric_type = ""
    grouped["FABRIC TYPE"] = fabric_type if fabric_type else "Knitted"

    # HS CODE and Country
    grouped["HS CODE"] = "61112000"
    grouped["COUNTRY OF ORIGIN"] = "India"

    # Ensure final ordering & presence of expected columns
    final_cols = [
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
    for c in final_cols:
        if c not in grouped.columns:
            # add defaults for any missing
            if c in ["QTY"]:
                grouped[c] = 0
            elif c in ["UNIT PRICE", "AMOUNT"]:
                grouped[c] = 0.0
            else:
                grouped[c] = ""
    grouped = grouped[final_cols]

    # Reset index and return
    grouped = grouped.reset_index(drop=True)
    return grouped


# ===== PDF Generator with ALL sections aligned to product table width =====
def generate_proforma_invoice(df, form_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=24, bottomMargin=24, leftMargin=24, rightMargin=24)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('CustomTitle', parent=styles['Normal'],
                                 fontSize=14, alignment=TA_CENTER,
                                 fontName='Helvetica-Bold', spaceAfter=6)
    header_style = ParagraphStyle('HeaderStyle', parent=styles['Normal'],
                                  fontSize=9, fontName='Helvetica-Bold', alignment=TA_LEFT)
    normal_style = ParagraphStyle('NormalStyle', parent=styles['Normal'],
                                  fontSize=8, alignment=TA_LEFT)

    elements.append(Paragraph("PROFORMA INVOICE", title_style))

    # CRITICAL: Define the SAME total width for ALL tables to match the product table
    # Product table column widths (must match exactly)
    product_col_widths = [0.8 * inch, 1.3 * inch, 0.8 * inch, 0.7 * inch,
                          1.1 * inch, 0.7 * inch, 0.5 * inch, 0.6 * inch, 0.8 * inch]
    total_table_width = sum(product_col_widths)  # This is our reference width
    
    # All header sections will use this same total width split into 2 columns
    header_col_widths = [total_table_width / 2, total_table_width / 2]

    # Supplier + PI block
    supplier_data = [
        [Paragraph("<b>Supplier Name:</b>", header_style),
         Paragraph(f"<b>No. & date of PI:</b> {form_data['pi_number']}", header_style)],
        [Paragraph("<b>SAR APPARELS INDIA PVT.LTD.</b>", header_style), ""],
        ["", Paragraph(f"<b>Landmark order Reference:</b> {form_data['order_ref']}", normal_style)],
        [Paragraph("<b>Address:</b> 6, Picaso Bithi, Kolkata - 700017", normal_style),
         Paragraph(f"<b>Buyer Name:</b> {form_data['buyer_name']}", normal_style)],
        [Paragraph("<b>Phone:</b> 9817473373", normal_style),
         Paragraph(f"<b>Brand Name:</b> {form_data['brand_name']}", normal_style)],
        [Paragraph("<b>Fax:</b> N.A.", normal_style), ""]
    ]
    elements.append(Table(supplier_data, colWidths=header_col_widths,
                          style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                 ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                                 ('VALIGN', (0, 0), (-1, -1), 'TOP')]))

    # Consignee + Bank block
    consignee_data = [
        [Paragraph("<b>Consignee:</b>", header_style),
         Paragraph(f"<b>Payment Term:</b> {form_data['payment_term']}", normal_style)],
        [Paragraph(form_data['consignee_name'], normal_style), ""],
        [Paragraph(form_data['consignee_address'], normal_style),
         Paragraph("<b>Bank Details (Including Swift/IBAN)</b>", header_style)],
        [Paragraph(form_data['consignee_tel'], normal_style), ""],
        ["", Paragraph(f"<b>Beneficiary</b> :- {form_data['bank_beneficiary']}", normal_style)],
        ["", Paragraph(f"<b>Account No</b> :- {form_data['bank_account']}", normal_style)],
        ["", Paragraph(f"<b>BANK'S NAME</b> :- {form_data['bank_name']}", normal_style)],
        ["", Paragraph(f"<b>BANK ADDRESS</b> :- {form_data['bank_address']}", normal_style)],
        ["", Paragraph(f"<b>SWIFT CODE</b> :- {form_data['bank_swift']}", normal_style)],
        ["", Paragraph(f"<b>BANK CODE</b> :- {form_data['bank_code']}", normal_style)]
    ]
    elements.append(Table(consignee_data, colWidths=header_col_widths,
                          style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                 ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                                 ('VALIGN', (0, 0), (-1, -1), 'TOP')]))

    # Shipping + Remarks block
    shipping_data = [
        [Paragraph(f"<b>Loading Country:</b> {form_data['loading_country']}", normal_style),
         Paragraph("<b>L/C Advising Bank:</b> (If applicable)", normal_style)],
        [Paragraph(f"<b>Port of Loading:</b> {form_data['port_loading']}", normal_style), ""],
        [Paragraph(f"<b>Agreed Shipment Date:</b> {form_data['shipment_date']}", normal_style), ""],
        [Paragraph(f"<b>Remarks:</b> {form_data['remarks']}", normal_style), ""]
    ]
    elements.append(Table(shipping_data, colWidths=header_col_widths,
                          style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                 ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                                 ('VALIGN', (0, 0), (-1, -1), 'TOP')]))

    # Goods + Currency block - uses same total width but different proportions
    goods_data = [
        [Paragraph(f"<b>Description of goods:</b> {form_data['goods_desc']}", normal_style),
         Paragraph("<b>CURRENCY: USD</b>", ParagraphStyle('RightAlign',
                                                          parent=normal_style, alignment=TA_RIGHT,
                                                          fontName='Helvetica-Bold'))]
    ]
    elements.append(Table(goods_data, colWidths=[total_table_width * 0.75, total_table_width * 0.25],
                          style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                 ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
                                 ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')]))

    # === Product table (using the exact same widths as defined above)
    table_headers = ["STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE\nKNITTED / WOVEN",
                     "H.S NO\n(8digit)", "COMPOSITION OF\nMATERIAL", "COUNTRY\nOF\nORIGIN",
                     "QTY", "UNIT\nPRICE\nFOB", "AMOUNT"]
    table_data = [table_headers]

    total_qty = 0
    total_amount = 0.0
    for _, row in df.iterrows():
        qty = int(row.get("QTY", 0) or 0)
        unit_price = float(row.get("UNIT PRICE", 0.0) or 0.0)
        amount = float(row.get("AMOUNT", qty * unit_price) or (qty * unit_price))
        total_qty += qty
        total_amount += amount

        table_data.append([
            str(row.get("STYLE NO", "")),
            str(row.get("ITEM DESCRIPTION", "")),
            str(row.get("FABRIC TYPE", "")),
            str(row.get("HS CODE", "")),
            str(row.get("COMPOSITION", "")),
            str(row.get("COUNTRY OF ORIGIN", "")),
            f"{qty:,}",
            f"{unit_price:.2f}",
            f"{amount:.2f}",
        ])

    # total row
    table_data.append(["", "", "", "", "", "TOTAL", f"{total_qty:,}", "", f"USD {total_amount:.2f}"])

    product_table = Table(table_data, colWidths=product_col_widths)
    product_table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
        ('BOX', (0, 0), (-1, -1), 1, colors.black)
    ]))
    elements.append(product_table)

    # Signature section - NOW aligned to same total width as product table
    signature_data = [
        ["Signed by …………………….(Affix Stamp here)",
         "for RNA Resources Group Ltd-Landmark (Babyshop)"],
        ["Terms & Conditions (If Any)", ""]
    ]
    signature_table = Table(signature_data, colWidths=header_col_widths)  # Same width as header sections
    signature_table.setStyle(TableStyle([
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
        ('BOX', (0, 0), (-1, -1), 1, colors.black),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
    ]))
    elements.append(signature_table)

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

        with st.form("invoice_form"):
            st.subheader("✍️ Enter Invoice Details")
            pi_number = st.text_input("PI No. & Date", "SAR/LG/XXXX Dt. 04/09/2025")
            order_ref = st.text_input("Landmark order Reference", "CPO/47062/25")
            buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
            brand_name = st.text_input("Brand Name", "Juniors")
            consignee_name = st.text_input("Consignee Name", "RNA Resource Group Ltd - Landmark (Babyshop)")
            consignee_address = st.text_area("Consignee Address", "P.O Box 25030, Dubai, UAE")
            consignee_tel = st.text_input("Consignee Tel/Fax", "Tel: 00971 4 8095500, Fax: 00971 4 8095555/66")
            payment_term = st.text_input("Payment Term", "T/T")
            bank_beneficiary = st.text_input("Bank Beneficiary", "SAR APPARELS INDIA PVT.LTD.")
            bank_account = st.text_input("Account No", "2112819952")
            bank_name = st.text_input("Bank Name", "KOTAK MAHINDRA BANK LTD")
            bank_address = st.text_area("Bank Address", "2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR, KOLKATA-700001")
            bank_swift = st.text_input("SWIFT", "KKBKINBBCPC")
            bank_code = st.text_input("Bank Code", "0323")
            loading_country = st.text_input("Loading Country", "India")
            port_loading = st.text_input("Port of Loading", "Mumbai")
            shipment_date = st.text_input("Agreed Shipment Date", "07/02/2025")
            remarks = st.text_area("Remarks", "")
            goods_desc = st.text_input("Description of goods", "Value Packs")
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            form_data = {
                "pi_number": pi_number,
                "order_ref": order_ref,
                "buyer_name": buyer_name,
                "brand_name": brand_name,
                "consignee_name": consignee_name,
                "consignee_address": consignee_address,
                "consignee_tel": consignee_tel,
                "payment_term": payment_term,
                "bank_beneficiary": bank_beneficiary,
                "bank_account": bank_account,
                "bank_name": bank_name,
                "bank_address": bank_address,
                "bank_swift": bank_swift,
                "bank_code": bank_code,
                "loading_country": loading_country,
                "port_loading": port_loading,
                "shipment_date": shipment_date,
                "remarks": remarks,
                "goods_desc": goods_desc,
            }
            pdf_buffer = generate_proforma_invoice(df, form_data)
            st.success("✅ PDF Generated Successfully!")
            st.download_button("⬇️ Download Proforma Invoice", data=pdf_buffer,
                               file_name="Proforma_Invoice.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"❌ Error: {e}")
