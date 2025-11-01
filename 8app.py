import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
import io

# -------------------------------
# Custom UI Styling (Questrial + Color Palette)
# -------------------------------
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Questrial&display=swap');

    html, body, [class*="css"] {
        font-family: 'Questrial', sans-serif;
        color: #FFFFFF;
        background-color: #000000;
    }

    .stApp {
        background-color: #000000;
        color: white;
    }

    section[data-testid="stSidebar"] {
        background-color: #0D0D0D;
        border-right: 1px solid rgba(255,255,255,0.1);
    }

    h1, h2, h3, h4, h5, h6 {
        font-family: 'Questrial', sans-serif !important;
        color: white !important;
    }

    div.stButton > button {
        background: linear-gradient(90deg, #FF6F61 0%, #FFD300 100%);
        color: black;
        border: none;
        border-radius: 20px;
        padding: 0.6rem 1.5rem;
        font-size: 16px;
        transition: 0.3s ease;
        font-weight: 500;
    }

    div.stButton > button:hover {
        background: linear-gradient(90deg, #FFD300 0%, #FF6F61 100%);
        color: white;
        transform: scale(1.03);
    }

    div.stButton > button:active {
        background: #FFD300 !important;
        color: black !important;
    }

    .stTextInput > div > div > input, 
    .stFileUploader label div div {
        background-color: #111111;
        color: white;
        border: 1px solid rgba(255,255,255,0.1);
        border-radius: 10px;
    }

    .stTabs [role="tablist"] {
        background-color: #111111;
        border-radius: 15px;
        padding: 6px;
    }

    .stTabs [role="tab"] {
        color: #CCCCCC;
        border-radius: 12px;
        font-size: 16px;
        transition: all 0.3s;
    }

    .stTabs [aria-selected="true"] {
        background-color: #66CCFF !important;
        color: black !important;
        font-weight: bold;
    }

    .stDataFrame, .stTable {
        background-color: #111111;
        border-radius: 12px;
        border: 1px solid rgba(255,255,255,0.1);
    }

    .streamlit-expanderHeader {
        background-color: #0D0D0D !important;
        color: #FFD300 !important;
        font-weight: 500;
    }

    .stDownloadButton button {
        background-color: #66CCFF;
        color: black;
        border-radius: 20px;
        padding: 0.5rem 1.2rem;
        border: none;
        transition: 0.3s ease;
    }

    .stDownloadButton button:hover {
        background-color: #FFD300;
        color: black;
    }

    </style>
""", unsafe_allow_html=True)

# -------------------------------
# App Title and Description
# -------------------------------
st.title("🧾 Automated Invoice Generator")
st.write("Upload your Excel file, review your data, and generate styled PDF invoices instantly!")

# -------------------------------
# File Upload
# -------------------------------
uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success("✅ File uploaded successfully!")
    st.write("Preview of your data:")
    st.dataframe(df)

    # -------------------------------
    # Invoice Generation
    # -------------------------------
    def indian_format(number):
        x = str(int(number))
        if len(x) <= 3:
            return x
        else:
            return x[-3:] + "," + ",".join(
                [x[max(i - 2, 0):i] for i in range(len(x) - 3, 0, -2)]
            )[::-1]

    def create_invoice(data):
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=A4)
        elements = []

        styles = getSampleStyleSheet()
        styles.add(ParagraphStyle(name="Center", alignment=TA_CENTER))
        styles.add(ParagraphStyle(name="Right", alignment=TA_RIGHT))
        styles.add(ParagraphStyle(name="Left", alignment=TA_LEFT))

        elements.append(Paragraph("INVOICE", styles["Center"]))
        elements.append(Spacer(1, 12))

        table_data = [["Item", "Quantity", "Price", "Total"]]
        for i in range(len(data)):
            table_data.append([
                data["Item"][i],
                data["Quantity"][i],
                data["Price"][i],
                data["Quantity"][i] * data["Price"][i],
            ])

        t = Table(table_data)
        t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#66CCFF")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
            ('GRID', (0, 0), (-1, -1), 1, colors.grey),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ]))
        elements.append(t)

        total_amount = sum(data["Quantity"] * data["Price"])
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f"<b>Total:</b> ₹{indian_format(total_amount)}", styles["Right"]))

        doc.build(elements)
        pdf = buffer.getvalue()
        buffer.close()
        return pdf

    if st.button("Generate Invoice"):
        pdf_data = create_invoice(df)
        st.success("✅ Invoice generated successfully!")

        st.download_button(
            label="⬇️ Download Invoice PDF",
            data=pdf_data,
            file_name="invoice.pdf",
            mime="application/pdf",
        )
else:
    st.info("Please upload an Excel file to get started.")
