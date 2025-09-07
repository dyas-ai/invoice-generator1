import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph
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
            df_raw.iloc[stacked_header_idx].astype(str).fillna("") + " " + df_raw.iloc[header_row_idx].astype(str).fillna("")
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
    doc = SimpleDocTemplate(buffer, pagesize=A4, topMargin=24, bottomMargin=24, leftMargin=24, rightMargin=24)
    elements = []

    styles = getSampleStyleSheet()
    header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=7, fontName='Helvetica-Bold', alignment=TA_CENTER)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=6, alignment=TA_CENTER)

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

    # TOTAL row with merged cells
    table_data.append(["TOTAL", "", "", "", "", "", f"{total_qty:,}", "", f"USD {total_amount:.2f}"])

    product_col_widths = [0.8*inch, 1.3*inch, 0.8*inch, 0.7*inch, 1.1*inch, 0.7*inch, 0.5*inch, 0.6*inch, 0.8*inch]
    product_table = Table(table_data,colWidths=product_col_widths)

    product_table.setStyle(TableStyle([
        # No grey header background
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

        # Merge TOTAL row
        ('SPAN',(0,-1),(5,-1)),  # TOTAL across first 6 cols
        ('SPAN',(6,-1),(7,-1)),  # merge QTY + UNIT PRICE cols
        ('ALIGN',(0,-1),(5,-1),'CENTER'),
        ('ALIGN',(6,-1),(7,-1),'CENTER'),
        ('ALIGN',(8,-1),(8,-1),'CENTER'),
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
    ]))
    elements.append(product_table)

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
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            form_data = {"pi_number":pi_number}
            pdf_buffer = generate_proforma_invoice(df, form_data)
            st.success("✅ PDF Generated Successfully!")
            st.download_button("⬇️ Download Proforma Invoice", data=pdf_buffer,
                               file_name="Proforma_Invoice.pdf", mime="application/pdf")
    except Exception as e:
        st.error(f"❌ Error: {e}")
