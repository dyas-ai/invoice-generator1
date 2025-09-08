import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import io
from datetime import datetime
import num2words

# ---------- PDF GENERATION ----------
def generate_invoice(df, supplier_name, supplier_address, supplier_phone, supplier_fax,
                     pi_number, pi_date, landmark_ref, buyer_name, brand_name,
                     loading_country, port_loading, shipment_date,
                     total_amount, stamp_url):

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=30, leftMargin=30,
                            topMargin=30, bottomMargin=18)

    elements = []
    styles = getSampleStyleSheet()
    normal = styles["Normal"]

    # -------- Supplier Section --------
    supplier_info = [
        [Paragraph(f"<b>{supplier_name}</b>", normal)],
        [Paragraph(supplier_address, normal)],
        [Paragraph(f"Phone: {supplier_phone}", normal)],
        [Paragraph(f"Fax: {supplier_fax}", normal)],
    ]
    supplier_table = Table(supplier_info, colWidths=[450])
    supplier_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('BOTTOMPADDING', (0, 2), (0, 2), 0),  # tighten Phone/Fax gap
    ]))
    elements.append(supplier_table)
    elements.append(Spacer(1, 12))

    # -------- PI Number & Date --------
    pi_data = [
        [Paragraph(f"<b>PI Number:</b> {pi_number}", normal),
         Paragraph(f"<b>Date:</b> {pi_date}", normal)]
    ]
    pi_table = Table(pi_data, colWidths=[225, 225])
    pi_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.black),
    ]))
    elements.append(pi_table)
    elements.append(Spacer(1, 6))

    # -------- Landmark Section --------
    landmark_data = [
        [Paragraph(f"<b>Landmark Order Reference:</b> {landmark_ref}", normal)],
        [Paragraph(f"<b>Buyer Name:</b> {buyer_name}", normal)],
        [Paragraph(f"<b>Brand Name:</b> {brand_name}", normal)],
    ]
    landmark_table = Table(landmark_data, colWidths=[450])
    landmark_table.setStyle(TableStyle([
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
    ]))
    elements.append(landmark_table)
    elements.append(Spacer(1, 12))

    # -------- Loading / Shipment --------
    load_data = [
        [Paragraph(f"<b>Loading Country:</b> {loading_country}", normal)],
        [Paragraph(f"<b>Port of Loading:</b> {port_loading}", normal)],
        [Paragraph(f"<b>Agreed Shipment Date:</b> {shipment_date}", normal)],
    ]
    load_table = Table(load_data, colWidths=[450])
    load_table.setStyle(TableStyle([
        ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
    ]))
    elements.append(load_table)
    elements.append(Spacer(1, 12))

    # -------- Goods Table --------
    table_data = [df.columns.tolist()] + df.values.tolist()

    # Add description + currency rows
    table_data.append(["Description of goods: Value Packs", "", "", ""])
    table_data.append(["CURRENCY: USD", "", "", ""])

    t = Table(table_data, colWidths=[200, 80, 80, 90])

    # Custom row heights for description + currency
    row_heights = [18] * len(table_data)
    row_heights[-2] = 40  # taller Description row
    row_heights[-1] = 30  # taller Currency row
    t._argH = row_heights

    t.setStyle(TableStyle([
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
        ('VALIGN', (0, -2), (0, -2), 'MIDDLE'),
        ('VALIGN', (0, -1), (0, -1), 'MIDDLE'),
        ('SPAN', (0, -2), (-1, -2)),  # description across all
        ('SPAN', (0, -1), (-1, -1)),  # currency across all
    ]))
    elements.append(t)
    elements.append(Spacer(1, 12))

    # -------- Total Section --------
    total_in_words = num2words.num2words(total_amount, to='currency', lang='en').upper()
    total_in_words = f"TOTAL IN WORDS: USD {total_in_words.replace('EUROS', '').strip()}"

    totals_data = [
        [Paragraph(total_in_words, normal), f"TOTAL: USD {total_amount:,.2f}"],
    ]
    totals_table = Table(totals_data, colWidths=[300, 150])
    totals_table.setStyle(TableStyle([
        ('SPAN', (0, 0), (0, 0)),
        ('ALIGN', (0, 0), (0, 0), 'LEFT'),
        ('ALIGN', (1, 0), (1, 0), 'RIGHT'),
    ]))
    elements.append(totals_table)
    elements.append(Spacer(1, 12))

    # -------- Terms & Conditions --------
    terms = Paragraph("<b>Terms & Conditions:</b><br/>All disputes subject to jurisdiction.", normal)
    elements.append(terms)
    elements.append(Spacer(1, 24))

    # -------- Signature Section --------
    sign_data = [
        [Paragraph("<b>For Supplier</b>", normal)],
    ]
    sign_table = Table(sign_data, colWidths=[450], rowHeights=[20])
    sign_table.setStyle(TableStyle([
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ]))
    elements.append(sign_table)

    # Insert e-stamp image
    if stamp_url:
        try:
            img = Image(stamp_url, width=120, height=120)  # adjust size as needed
            elements.append(img)
        except Exception as e:
            elements.append(Paragraph(f"[Error loading stamp image: {e}]", normal))

    elements.append(Spacer(1, 40))
    elements.append(Paragraph("Authorised Signatory", normal))

    doc.build(elements)
    pdf = buffer.getvalue()
    buffer.close()
    return pdf

# ---------- STREAMLIT APP ----------
def main():
    st.title("Proforma Invoice Generator")

    supplier_name = st.text_input("Supplier Name", "ABC Textiles Pvt Ltd")
    supplier_address = st.text_area("Supplier Address", "123, Industrial Area, New Delhi")
    supplier_phone = st.text_input("Phone", "9817473373")
    supplier_fax = st.text_input("Fax", "N.A")

    pi_number = st.text_input("PI Number", "PI-001")
    pi_date = st.date_input("PI Date", datetime.today()).strftime("%d/%m/%Y")

    landmark_ref = st.text_input("Landmark Order Reference", "CPO/47062/25")
    buyer_name = st.text_input("Buyer Name", "LANDMARK GROUP")
    brand_name = st.text_input("Brand Name", "Juniors")

    loading_country = st.text_input("Loading Country", "India")
    port_loading = st.text_input("Port of Loading", "Mumbai")
    shipment_date = st.date_input("Agreed Shipment Date", datetime.today()).strftime("%d/%m/%Y")

    # File uploader for e-stamp (optional override)
    stamp_url = "https://raw.githubusercontent.com/dyas-ai/invoice-generator1/main/Screenshot%202025-09-06%20163303.png"

    # Table data
    st.subheader("Add Items")
    data = {
        "Description": ["Item 1", "Item 2"],
        "Quantity": [10, 20],
        "Unit Price": [100, 200],
        "Amount": [1000, 4000]
    }
    df = pd.DataFrame(data)

    total_amount = df["Amount"].sum()

    if st.button("Generate Invoice"):
        pdf = generate_invoice(df, supplier_name, supplier_address, supplier_phone, supplier_fax,
                               pi_number, pi_date, landmark_ref, buyer_name, brand_name,
                               loading_country, port_loading, shipment_date,
                               total_amount, stamp_url)

        st.download_button("Download Invoice", data=pdf,
                           file_name="invoice.pdf", mime="application/pdf")

if __name__ == "__main__":
    main()
