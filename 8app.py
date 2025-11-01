import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from io import BytesIO
from datetime import datetime

# =========================
# Streamlit Page Setup
# =========================
st.set_page_config(page_title="Invoice Generator", layout="wide")

# =========================
# Minimal Black & White Theme Styling
# =========================
st.markdown("""
<style>
/* Overall app */
.stApp {
    background-color: #ffffff;
    color: #000000;
    font-family: 'Inter', sans-serif;
}

/* Headings */
h1, h2, h3, h4, h5, h6 {
    color: #000000;
    font-weight: 600;
    letter-spacing: -0.3px;
}

/* Paragraph text */
p, label, span, div {
    color: #000000;
}

/* Rounded black buttons */
div.stButton > button:first-child {
    background-color: #000000;
    color: #ffffff;
    border: none;
    border-radius: 10px;
    padding: 0.6em 1.5em;
    font-weight: 500;
    letter-spacing: 0.3px;
    transition: all 0.2s ease;
}
div.stButton > button:first-child:hover {
    background-color: #222222;
    transform: scale(1.03);
}

/* Input fields */
.stTextInput > div > div > input,
.stNumberInput input,
.stTextArea textarea {
    border-radius: 8px !important;
    border: 1px solid #000000 !important;
    background-color: #ffffff !important;
    color: #000000 !important;
}

/* File uploader label */
.stFileUploader label {
    color: #000000 !important;
    font-weight: 500 !important;
}

/* Tables */
.stDataFrame, .stTable {
    border-radius: 10px;
    border: 1px solid #00000022;
    padding: 8px;
}

/* Divider lines */
hr {
    border: 0;
    border-top: 1px solid #00000022;
    margin: 1.5rem 0;
}

/* Remove Streamlit watermark and extra menus */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)

# =========================
# App Title
# =========================
st.title("🧾 Invoice Generator")
st.markdown("---")

# =========================
# File Upload Section
# =========================
uploaded_file = st.file_uploader("📤 Upload Excel File", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)

    st.subheader("📄 Uploaded Data Preview")
    st.dataframe(df.head())

    # Column mapping input
    st.markdown("### 🧩 Column Mapping")
    col1, col2, col3 = st.columns(3)
    with col1:
        name_col = st.selectbox("Select Client Name Column", df.columns)
    with col2:
        amount_col = st.selectbox("Select Amount Column", df.columns)
    with col3:
        date_col = st.selectbox("Select Date Column", df.columns)

    st.markdown("---")

    # Invoice info
    st.subheader("🏷️ Invoice Details")
    colA, colB = st.columns(2)
    with colA:
        company_name = st.text_input("Company Name", "Your Company Pvt. Ltd.")
        invoice_prefix = st.text_input("Invoice Prefix", "INV")
    with colB:
        address = st.text_area("Company Address", "123 Street Name\nCity, State - ZIP")
        invoice_date = st.date_input("Invoice Date", datetime.today())

    st.markdown("---")

    # =========================
    # Generate Invoice PDFs
    # =========================
    if st.button("Generate Invoices"):
        buffer = BytesIO()

        for index, row in df.iterrows():
            pdf = BytesIO()
            doc = SimpleDocTemplate(pdf, pagesize=A4, rightMargin=30, leftMargin=30, topMargin=30, bottomMargin=18)
            story = []
            styles = getSampleStyleSheet()

            client_name = str(row[name_col])
            amount = row[amount_col]
            date_val = row[date_col]

            # Title
            story.append(Paragraph(f"<b>{company_name}</b>", styles['Title']))
            story.append(Paragraph(address.replace("\n", "<br/>"), styles['Normal']))
            story.append(Spacer(1, 12))
            story.append(Paragraph(f"<b>Invoice:</b> {invoice_prefix}-{index+1}", styles['Heading3']))
            story.append(Paragraph(f"<b>Date:</b> {invoice_date.strftime('%d %B %Y')}", styles['Normal']))
            story.append(Spacer(1, 12))

            # Client info
            story.append(Paragraph(f"<b>Bill To:</b> {client_name}", styles['Heading3']))
            story.append(Spacer(1, 12))

            # Table
            data = [["Description", "Amount (INR)"],
                    [f"Invoice for {client_name}", f"₹{amount:,.2f}"]]
            table = Table(data, colWidths=[300, 150])
            table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.black),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (1, 1), (-1, -1), 'RIGHT'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
                ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
                ('GRID', (0, 0), (-1, -1), 0.5, colors.black),
            ]))
            story.append(table)
            story.append(Spacer(1, 24))

            story.append(Paragraph(f"<b>Total Due:</b> ₹{amount:,.2f}", styles['Heading2']))
            story.append(Spacer(1, 12))
            story.append(Paragraph("Thank you for your business!", styles['Normal']))

            doc.build(story)
            pdf.seek(0)
            buffer.write(pdf.read())

        buffer.seek(0)
        st.download_button(
            label="⬇️ Download All Invoices (ZIP)",
            data=buffer,
            file_name=f"Invoices_{datetime.now().strftime('%Y%m%d')}.zip",
            mime="application/zip"
        )

else:
    st.info("Please upload an Excel file to begin.")
