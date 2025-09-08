import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT, TA_TOP
import io
from datetime import datetime
import num2words

def generate_proforma_invoice(df, form_data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4,
                            rightMargin=30, leftMargin=30,
                            topMargin=30, bottomMargin=18)
    elements = []
    styles = getSampleStyleSheet()
    normal_style = styles["Normal"]

    # Title
    title_style = ParagraphStyle("TitleStyle", parent=styles["Heading1"], alignment=TA_CENTER, fontSize=14, spaceAfter=10)
    elements.append(Paragraph("<b>PROFORMA INVOICE</b>", title_style))
    elements.append(Spacer(1, 6))

    # Supplier Info
    supplier_data = [
        ["Supplier Name & Address:", "ABC EXPORTS\n123 Street Name\nNew Delhi, India"],
        ["Phone:", "9817473373"],
        ["Fax:", "N.A."]
    ]
    supplier_table = Table(supplier_data, colWidths=[120, 350])
    supplier_table.setStyle(TableStyle([
        ("ALIGN", (0, 0), (-1, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(supplier_table)
    elements.append(Spacer(1, 8))

    # PI & Date
    pi_data = [
        ["PI No:", form_data["pi_number"], "Date:", datetime.today().strftime("%d/%m/%Y")]
    ]
    pi_table = Table(pi_data, colWidths=[50, 200, 40, 200])
    pi_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(pi_table)

    # Landmark Section
    landmark_data = [
        ["Landmark order Reference:", form_data["order_ref"]],
        ["Buyer Name:", form_data["buyer_name"]],
        ["Brand Name:", form_data["brand_name"]],
    ]
    landmark_table = Table(landmark_data, colWidths=[150, 380])
    landmark_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("VALIGN", (0, 0), (-1, -1), TA_TOP),
    ]))
    elements.append(landmark_table)
    elements.append(Spacer(1, 8))

    # Loading Country / Port / Shipment Date
    loading_data = [
        ["Loading Country:", form_data["loading_country"]],
        ["Port of Loading:", form_data["port_loading"]],
        ["Agreed Shipment Date:", form_data["shipment_date"]],
    ]
    loading_table = Table(loading_data, colWidths=[150, 380])
    loading_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
        ("TOPPADDING", (0, 0), (-1, -1), 2),
    ]))
    elements.append(loading_table)
    elements.append(Spacer(1, 12))

    # Goods Table
    data = [df.columns.tolist()] + df.values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(table)
    elements.append(Spacer(1, 12))

    # Description of Goods (taller row)
    total_table_width = 530
    goods_data = [[Paragraph(f"<b>Description of goods:</b> {form_data['goods_desc']}",
                              ParagraphStyle('Goods', parent=normal_style, fontSize=7)), ""]]
    goods_table = Table(goods_data, colWidths=[total_table_width, 0],
                        style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                               ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])
    goods_table._argH[0] = 40  # taller row
    elements.append(goods_table)

    # Currency (taller row)
    currency_data = [["",
                      Paragraph("<b>CURRENCY: USD</b>",
                                ParagraphStyle('Currency', parent=normal_style,
                                               fontSize=8, alignment=TA_RIGHT, fontName='Helvetica-Bold'))]]
    currency_table = Table(currency_data, colWidths=[total_table_width*0.75, total_table_width*0.25],
                           style=[('BOX', (0, 0), (-1, -1), 1, colors.black),
                                  ('VALIGN', (0, 0), (-1, -1), 'MIDDLE')])
    currency_table._argH[0] = 40  # taller row
    elements.append(currency_table)
    elements.append(Spacer(1, 12))

    # Totals
    total_qty = df["Quantity"].sum()
    total_value = df["Value"].sum()
    totals_data = [["TOTAL", total_qty, f"USD {total_value:,.2f}"]]
    totals_table = Table(totals_data, colWidths=[300, 100, 130])
    totals_table.setStyle(TableStyle([
        ("GRID", (0, 0), (-1, -1), 1, colors.black),
        ("BACKGROUND", (0, 0), (-1, -1), colors.lightgrey),
        ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (1, 0), (-1, -1), "CENTER"),
    ]))
    elements.append(totals_table)
    elements.append(Spacer(1, 8))

    # Total in words (single line, spanning columns)
    total_in_words = num2words.num2words(total_value, to="currency", lang="en", currency="USD").upper()
    total_words_para = Paragraph(f"<b>TOTAL IN WORDS:</b> USD {total_in_words}",
                                 ParagraphStyle('TotalWords', parent=normal_style,
                                                fontSize=8, alignment=TA_LEFT, fontName="Helvetica-Bold"))
    total_words_table = Table([[total_words_para]], colWidths=[530])
    total_words_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
    ]))
    elements.append(total_words_table)
    elements.append(Spacer(1, 12))

    # Terms and Conditions
    terms_para = Paragraph("<b>Terms & Conditions:</b><br/>" + form_data["remarks"],
                           ParagraphStyle('Terms', parent=normal_style, fontSize=8, alignment=TA_LEFT))
    elements.append(terms_para)
    elements.append(Spacer(1, 20))

    # Signature section with e-stamp
    img = Image("https://raw.githubusercontent.com/dyas-ai/invoice-generator1/main/Screenshot%202025-09-06%20163303.png",
                width=120, height=60)
    signature_data = [
        ["", ""],
        [img, ""],
        ["Sign Here", ""]
    ]
    signature_table = Table(signature_data, colWidths=[265, 265])
    signature_table.setStyle(TableStyle([
        ("BOX", (0, 0), (-1, -1), 1, colors.black),
        ("ALIGN", (0, 1), (0, 1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
    ]))
    elements.append(signature_table)

    doc.build(elements)
    buffer.seek(0)
    return buffer

def main():
    st.title("📄 Proforma Invoice Generator")

    st.write("Upload your order data CSV:")
    uploaded_file = st.file_uploader("Choose a file", type=["csv"])

    if uploaded_file is not None:
        df = pd.read_csv(uploaded_file)
        st.write("### Preview Data")
        st.dataframe(df)

        with st.form("invoice_form"):
            pi_number = st.text_input("PI Number")
            order_ref = st.text_input("Landmark Order Reference")
            buyer_name = st.text_input("Buyer Name")
            brand_name = st.text_input("Brand Name")
            consignee_name = st.text_input("Consignee Name")
            consignee_address = st.text_area("Consignee Address")
            consignee_tel = st.text_input("Consignee Telephone")
            payment_term = st.text_input("Payment Terms")
            bank_beneficiary = st.text_input("Bank Beneficiary")
            bank_account = st.text_input("Bank Account")
            bank_name = st.text_input("Bank Name")
            bank_address = st.text_area("Bank Address")
            bank_swift = st.text_input("Bank Swift Code")
            bank_code = st.text_input("Bank Code")
            loading_country = st.text_input("Loading Country")
            port_loading = st.text_input("Port of Loading")
            shipment_date = st.text_input("Agreed Shipment Date")
            remarks = st.text_area("Remarks / Terms & Conditions")
            goods_desc = st.text_area("Description of Goods")

            submitted = st.form_submit_button("Generate Proforma Invoice")

        try:
            if submitted:
                form_data = {
                    "pi_number": pi_number, "order_ref": order_ref, "buyer_name": buyer_name, "brand_name": brand_name,
                    "consignee_name": consignee_name, "consignee_address": consignee_address, "consignee_tel": consignee_tel,
                    "payment_term": payment_term, "bank_beneficiary": bank_beneficiary, "bank_account": bank_account,
                    "bank_name": bank_name, "bank_address": bank_address, "bank_swift": bank_swift, "bank_code": bank_code,
                    "loading_country": loading_country, "port_loading": port_loading, "shipment_date": shipment_date,
                    "remarks": remarks, "goods_desc": goods_desc
                }
                pdf_buffer = generate_proforma_invoice(df, form_data)
                st.success("✅ PDF Generated Successfully!")
                st.download_button("⬇️ Download Proforma Invoice", data=pdf_buffer,
                                   file_name="Proforma_Invoice.pdf", mime="application/pdf")
        except Exception as e:
            st.error(f"❌ Error: {e}")

if __name__ == "__main__":
    main()
