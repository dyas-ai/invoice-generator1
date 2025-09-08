import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib.units import inch
import io
from num2words import num2words

# ===== Preprocessing Function =====
def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    df_raw = pd.read_excel(uploaded_file, header=None)

    # detect header row
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

    # combine stacked headers
    if stacked_header_idx >= 0:
        headers = (
            df_raw.iloc[stacked_header_idx].astype(str).fillna("") + " " +
            df_raw.iloc[header_row_idx].astype(str).fillna("")
        )
    else:
        headers = df_raw.iloc[header_row_idx].astype(str).fillna("")
    headers = headers.str.strip().astype(str)

    # column mapping
    col_map = {
        "STYLE NO": ["Style", "Style No", "Item Style", "STYLE"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc", "DESC"],
        "COMPOSITION": ["Composition", "Fabric Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD", "Fob $", "Unit Price", "FOB"],
        "QTY": ["Total Qty", "Quantity", "Qty", "QTY"],
        "AMOUNT": ["Total Value", "Amount", "Value", "TOTAL VALUE"],
    }

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
        df_columns[target_col] = found

    # build dataframe
    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    if rename_dict:
        df = df.rename(columns=rename_dict)

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

    df["STYLE NO"] = df["STYLE NO"].astype(str).str.strip()
    df = df[~df["STYLE NO"].isin(["", "nan", "NaN", "None", "NONE"])]
    df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]

    df["QTY"] = pd.to_numeric(df.get("QTY", 0), errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df.get("UNIT PRICE", 0.0), errors="coerce").fillna(0.0).astype(float)
    df["AMOUNT"] = df["QTY"] * df["UNIT PRICE"]

    df = df[~((df["QTY"] == 0) & (df["UNIT PRICE"] == 0) & (df["STYLE NO"].str.strip() == ""))]

    group_by_cols = ["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"]
    for c in group_by_cols:
        if c not in df.columns:
            df[c] = "" if c != "UNIT PRICE" else 0.0

    grouped = (
        df.groupby(group_by_cols, dropna=False, as_index=False)
        .agg({"QTY": "sum"})
        .reset_index(drop=True)
    )
    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

    # static extras
    grouped["FABRIC TYPE"] = "Knitted"
    grouped["HS CODE"] = "61112000"
    grouped["COUNTRY OF ORIGIN"] = "India"

    final_cols = [
        "STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
        "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "AMOUNT",
    ]
    for c in final_cols:
        if c not in grouped.columns:
            grouped[c] = "" if c not in ["QTY", "UNIT PRICE", "AMOUNT"] else 0.0
    grouped = grouped[final_cols].reset_index(drop=True)
    return grouped

# ===== PDF Generator =====
def generate_proforma_invoice(df, form_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            topMargin=24, bottomMargin=24,
                            leftMargin=24, rightMargin=24)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontSize=12,
                                 alignment=TA_CENTER, fontName='Helvetica-Bold', spaceAfter=6)
    header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=7,
                                  fontName='Helvetica-Bold', alignment=TA_LEFT, 
                                  spaceBefore=0, spaceAfter=0, leading=8)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=6, alignment=TA_LEFT,
                                  spaceBefore=0, spaceAfter=0, leading=7)

    elements.append(Paragraph("PROFORMA INVOICE", title_style))

    # width setup
    product_col_widths = [0.8*inch, 1.3*inch, 0.8*inch, 0.7*inch,
                          1.1*inch, 0.7*inch, 0.5*inch, 0.6*inch, 0.8*inch]
    total_table_width = sum(product_col_widths)
    header_col_widths = [total_table_width/2, total_table_width/2]

    # Supplier section
    supplier_data = [
        [Paragraph("<b>Supplier Name:</b>", header_style),
         Paragraph(f"<b>No. & date of PI:</b> {form_data['pi_number']}", header_style)],
        [Paragraph("<b>SAR APPARELS INDIA PVT.LTD.</b><br/><b>Address:</b> 6, Picaso Bithi, Kolkata - 700017<br/><b>Phone:</b> 9817473373<br/><b>Fax:</b> N.A.", ParagraphStyle('SupplierDetail', parent=header_style, leading=6)),
         Paragraph(f"<b>Landmark order Reference:</b> {form_data['order_ref']}<br/><b>Buyer Name:</b> {form_data['buyer_name']}<br/><b>Brand Name:</b> {form_data['brand_name']}", ParagraphStyle('TopAlign', parent=header_style, alignment=TA_LEFT, spaceBefore=0))],
    ]
    elements.append(Table(supplier_data, colWidths=header_col_widths,
                          style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                 ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                                 ('LINEBELOW',(1,0),(1,0),1,colors.black),
                                 ('VALIGN',(0,1),(1,1),'TOP'),
                                 ('BOTTOMPADDING',(0,1),(0,1),6),
                                 ('BOTTOMPADDING',(1,1),(1,1),6)]))

    # Consignee section - ULTRA TIGHT SPACING
    # Create compact styles for bank details
    bank_style = ParagraphStyle('BankCompact', parent=normal_style, fontSize=7, fontName='Helvetica-Bold', 
                               leading=8, spaceAfter=0, spaceBefore=0, leftIndent=0, rightIndent=0)
    
    consignee_data = [
        [Paragraph("<b>Consignee:</b>", header_style),
         Paragraph(f"<b>Payment Term:</b> {form_data['payment_term']}", normal_style)],
        [Paragraph(f"{form_data['consignee_name']}<br/>{form_data['consignee_address']}<br/>{form_data['consignee_tel']}", 
                   ParagraphStyle('ConsigneeCompact', parent=normal_style, leading=8, spaceAfter=0, spaceBefore=0)),
         Paragraph(f"<br/><br/><b>Bank Details</b><br/><b>BENEFICIARY</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_beneficiary']}<br/><b>ACCOUNT NO</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_account']}<br/><b>BANK'S NAME</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_name']}<br/><b>BANK ADDRESS</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- 2 BRABOURNE ROAD, GOVIND BHAVAN,<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GROUND FLOOR, KOLKATA-700001<br/><b>SWIFT CODE</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_swift']}<br/><b>BANK CODE</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_code']}", 
                   bank_style)]
    ]
    consignee_table = Table(consignee_data, colWidths=header_col_widths,
                            style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                   ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                                   ('VALIGN',(0,0),(-1,-1),'TOP'),
                                   # Ultra tight spacing
                                   ('TOPPADDING',(0,0),(-1,-1),0),    # Zero top padding for all cells
                                   ('BOTTOMPADDING',(0,0),(-1,-1),1), # Minimal bottom padding for all cells
                                   ('LEFTPADDING',(0,0),(-1,-1),2),   # Minimal left padding
                                   ('RIGHTPADDING',(0,0),(-1,-1),2)]) # Minimal right padding

    elements.append(consignee_table)

    # Shipping section - REDUCED SPACING BELOW AGREED SHIPMENT DATE
    shipping_data = [
        [Paragraph(f"<b>Loading Country:</b> {form_data['loading_country']}", normal_style),
         Paragraph("<b>L/C Advising Bank:</b> (If applicable)", normal_style)],
        [Paragraph(f"<b>Port of Loading:</b> {form_data['port_loading']}", normal_style), ""],
        [Paragraph(f"<b>Agreed Shipment Date:</b> {form_data['shipment_date']}", normal_style), ""],
        ["", Paragraph(f"<b>Remarks:</b> {form_data['remarks']}", normal_style)]
    ]
    shipping_table = Table(shipping_data, colWidths=header_col_widths,
                          style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                 ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                                 ('VALIGN',(0,0),(-1,-1),'TOP'),
                                 ('TOPPADDING',(0,0),(-1,-1),1),    # Minimal top padding
                                 ('BOTTOMPADDING',(0,0),(-1,-1),1)]) # Minimal bottom padding
    
    # Set specific row heights to reduce spacing - reduced by 9 units total
    shipping_table._argH[0] = 9   # Loading Country row (was 18, now 9)
    shipping_table._argH[1] = 9   # Port of Loading row (was 18, now 9)
    shipping_table._argH[2] = 9   # Agreed Shipment Date row (was 18, now 9)
    shipping_table._argH[3] = 11  # Remarks row (was 20, now 11)
    
    elements.append(shipping_table)

    # Combined Goods and Currency block (NO LINE BETWEEN ROWS)
    combined_data = [
        # Row 1: Description of goods (left), empty right
        [Paragraph(f"<b>Description of goods:</b> {form_data['goods_desc']}", 
                   ParagraphStyle('Goods', parent=normal_style, fontSize=7)), ""],
        # Row 2: Empty left, Currency on right
        ["", Paragraph("<b>CURRENCY: USD</b>", 
                       ParagraphStyle('Currency', parent=normal_style, 
                                      fontSize=8, alignment=TA_RIGHT, fontName='Helvetica-Bold'))]
    ]
    
    combined_table = Table(combined_data, colWidths=header_col_widths,
                           style=[
                               # Outer border only - NO line between rows
                               ('BOX',(0,0),(-1,-1),1,colors.black),
                               ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                               ('VALIGN',(0,0),(-1,-1),'MIDDLE')
                           ])
    # Set both row heights to 25 units each
    combined_table._argH[0] = 25  # Row 1 height
    combined_table._argH[1] = 25  # Row 2 height
    elements.append(combined_table)

    # Product Table
    headers = ["STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE\nKNITTED / WOVEN","H.S NO\n(8digit)",
               "COMPOSITION OF\nMATERIAL","COUNTRY\nOF\nORIGIN","QTY","UNIT\nPRICE\nFOB","AMOUNT"]
    table_data = [headers]

    total_qty,total_amount = 0,0.0
    for _,row in df.iterrows():
        qty = int(row.get("QTY",0) or 0); price = float(row.get("UNIT PRICE",0.0) or 0.0)
        amt = float(row.get("AMOUNT", qty*price) or (qty*price))
        total_qty += qty; total_amount += amt
        table_data.append([str(row.get("STYLE NO","")),str(row.get("ITEM DESCRIPTION","")),
                           str(row.get("FABRIC TYPE","")),str(row.get("HS CODE","")),
                           str(row.get("COMPOSITION","")),str(row.get("COUNTRY OF ORIGIN","")),
                           f"{qty:,}",f"{price:.2f}",f"{amt:.2f}"])

    # TOTAL row
    table_data.append(
        ["TOTAL","","","","","",f"{total_qty:,}","",f"USD {total_amount:.2f}"]
    )

    product_table = Table(table_data,colWidths=product_col_widths, repeatRows=1)
    product_table.setStyle(TableStyle([
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),6),
        ('ALIGN',(0,0),(-1,-1),'CENTER'),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('BOX',(0,0),(-1,-1),1,colors.black),
        ('LINEBEFORE',(1,0),(1,-1),0.5,colors.black),
        ('LINEBEFORE',(2,0),(2,-1),0.5,colors.black),
        ('LINEBEFORE',(3,0),(3,-1),0.5,colors.black),
        ('LINEBEFORE',(4,0),(4,-1),0.5,colors.black),
        ('LINEBEFORE',(5,0),(5,-1),0.5,colors.black),
        ('LINEBEFORE',(6,0),(6,-1),0.5,colors.black),
        ('LINEBEFORE',(7,0),(7,-1),0.5,colors.black),
        ('LINEBEFORE',(8,0),(8,-1),0.5,colors.black),
        ('LINEBELOW',(0,0),(-1,0),0.5,colors.black),
        ('LINEABOVE',(0,-1),(-1,-1),0.5,colors.black),
        ('SPAN',(0,-1),(5,-1)),
        ('ALIGN',(0,-1),(5,-1),'CENTER'),
        ('SPAN',(6,-1),(7,-1)),
    ]))
    elements.append(product_table)

    # Signature block with e-stamp and total in words
    total_words_str = num2words(round(total_amount), to='cardinal', lang='en').upper()
    total_words_str = f"TOTAL IN WORDS: USD {total_words_str} DOLLARS"

    signature_data = [
        [Paragraph(total_words_str, ParagraphStyle('TotalWords', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, alignment=TA_LEFT)), ""],
        [Paragraph("Terms & Conditions (If Any)", normal_style), ""],
        [Image("https://raw.githubusercontent.com/dyas-ai/invoice-generator1/main/Screenshot%202025-09-06%20163303.png", width=2*inch, height=1*inch), ""],
        [Paragraph("Signed by …………………….(Affix Stamp here)", normal_style),
         Paragraph("for RNA Resources Group Ltd-Landmark (Babyshop)", normal_style)]
    ]
    signature_table = Table(signature_data, colWidths=header_col_widths,
                            style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                   ('VALIGN',(0,-1),(-1,-1),'BOTTOM'),
                                   ('SPAN',(0,0),(-1,0))])
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
        st.write("### Preview of Processed Data"); st.dataframe(df)

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
            form_data = {"pi_number":pi_number,"order_ref":order_ref,"buyer_name":buyer_name,"brand_name":brand_name,
                         "consignee_name":consignee_name,"consignee_address":consignee_address,"consignee_tel":consignee_tel,
                         "payment_term":payment_term,"bank_beneficiary":bank_beneficiary,"bank_account":bank_account,
                         "bank_name":bank_name,"bank_address":bank_address,"bank_swift":bank_swift,"bank_code":bank_code,
                         "loading_country":loading_country,"port_loading":port_loading,"shipment_date":shipment_date,
                         "remarks":remarks,"goods_desc":goods_desc}

            pdf_buffer = generate_proforma_invoice(df, form_data)
            st.download_button("📥 Download Proforma Invoice PDF", data=pdf_buffer, file_name="proforma_invoice.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"❌ Error: {e}")
