import streamlit as st
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from io import BytesIO
from datetime import date
import base64

# --- Page setup ---
st.set_page_config(page_title="Invoice Generator", page_icon="💼", layout="centered")

# --- Custom Styling ---
st.markdown(
    """
    <style>
    /* General page styling */
    body {
        background-color: #ffffff;
        color: #000000;
        font-family: 'Inter', sans-serif;
    }

    /* Streamlit default elements */
    .stApp {
        background-color: #ffffff;
        color: #000000;
    }

    /* Input fields and text boxes */
    .stTextInput > div > div > input {
        border-radius: 8px;
        border: 1px solid #000;
        background-color: #fff;
        color: #000;
    }

    .stNumberInput input {
        border-radius: 8px;
        border: 1px solid #000;
        background-color: #fff;
        color: #000;
    }

    /* Buttons */
    div.stButton > button {
        border-radius: 10px;
        background-color: #000000 !important;
        color: #ffffff !important;
        border: none;
        padding: 0.6em 1.2em;
        font-size: 1em;
        font-weight: 500;
        transition: all 0.3s ease;
    }

    div.stButton > button:hover {
        background-color: #333333 !important;
        color: #ffffff !important;
        transform: scale(1.02);
    }

    /* Headers and Titles */
    h1, h2, h3 {
        color: #000;
    }

    /* Horizontal line */
    hr {
        border: 1px solid #00000020;
        margin: 1.5em 0;
    }

    /* File uploader box */
    [data-testid="stFileUploader"] section {
        border: 1px solid #000;
        border-radius: 10px;
        background-color: #fff;
    }

    /* Expander aesthetics */
    details summary {
        font-weight: 500;
        border-radius: 8px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- App Title ---
st.title("💼 Invoice Generator")
st.caption("Generate professional black-and-white invoices with style.")

# --- Input Section ---
st.subheader("🧾 Invoice Details")
col1, col2 = st.columns(2)
with col1:
    invoice_no = st.text_input("Invoice Number", "INV-001")
    sender_name = st.text_input("Your Name / Company")
    sender_address = st.text_area("Your Address", height=90)
with col2:
    receiver_name = st.text_input("Client Name / Company")
    receiver_address = st.text_area("Client Address", height=90)
    invoice_date = st.date_input("Invoice Date", date.today())

st.write("---")

st.subheader("📦 Add Items")
items = []
num_items = st.number_input("Number of items", 1, 20, 1)
for i in range(int(num_items)):
    st.markdown(f"**Item {i+1}**")
    col1, col2, col3 = st.columns([3, 1, 1])
    with col1:
        item_name = st.text_input(f"Description {i+1}", key=f"name_{i}")
    with col2:
        qty = st.number_input(f"Qty {i+1}", min_value=1, key=f"qty_{i}")
    with col3:
        price = st.number_input(f"Price {i+1}", min_value=0.0, key=f"price_{i}")
    if item_name:
        items.append((item_name, qty, price))

st.write("---")

st.subheader("💰 Summary")
subtotal = sum(q * p for _, q, p in items)
tax_rate = st.number_input("Tax Rate (%)", 0.0, 50.0, 0.0)
tax_amount = subtotal * tax_rate / 100
total = subtotal + tax_amount

st.metric("Subtotal", f"₹{subtotal:,.2f}")
st.metric("Tax", f"₹{tax_amount:,.2f}")
st.metric("Total", f"₹{total:,.2f}")

st.write("---")

# --- Helper: Indian Format ---
def indian_format(amount):
    s, *d = str(round(amount, 2)).partition(".")
    r = ",".join([s[-i-2 if i else None:-i or None] for i in range(0, len(s), 2)][::-1])
    return "".join([r] + d)

# --- Generate PDF ---
def generate_invoice():
    buffer = BytesIO()
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    pdf.setTitle("Invoice")

    # Header
    pdf.setFont("Helvetica-Bold", 18)
    pdf.drawString(50, height - 50, "INVOICE")

    pdf.setFont("Helvetica", 10)
    pdf.drawString(50, height - 80, f"From: {sender_name}")
    pdf.drawString(50, height - 95, sender_address)
    pdf.drawString(50, height - 120, f"To: {receiver_name}")
    pdf.drawString(50, height - 135, receiver_address)
    pdf.drawString(50, height - 160, f"Date: {invoice_date}")
    pdf.drawString(400, height - 160, f"Invoice No: {invoice_no}")

    # Table header
    y = height - 200
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(50, y, "Description")
    pdf.drawString(300, y, "Qty")
    pdf.drawString(350, y, "Price")
    pdf.drawString(450, y, "Total")

    # Table content
    pdf.setFont("Helvetica", 10)
    y -= 20
    for item_name, qty, price in items:
        pdf.drawString(50, y, item_name)
        pdf.drawString(310, y, str(qty))
        pdf.drawString(360, y, f"₹{price:,.2f}")
        pdf.drawString(460, y, f"₹{qty*price:,.2f}")
        y -= 20

    # Totals
    y -= 20
    pdf.setFont("Helvetica-Bold", 10)
    pdf.drawString(400, y, "Subtotal:")
    pdf.drawString(460, y, f"₹{subtotal:,.2f}")
    y -= 15
    pdf.drawString(400, y, f"Tax ({tax_rate}%):")
    pdf.drawString(460, y, f"₹{tax_amount:,.2f}")
    y -= 15
    pdf.drawString(400, y, "Total:")
    pdf.drawString(460, y, f"₹{total:,.2f}")

    pdf.showPage()
    pdf.save()
    buffer.seek(0)
    return buffer

# --- Button and Download Section ---
if st.button("🧾 Generate Invoice"):
    pdf_buffer = generate_invoice()
    pdf_bytes = pdf_buffer.getvalue()
    b64 = base64.b64encode(pdf_bytes).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="invoice.pdf" style="text-decoration:none;"><button style="border-radius:10px;background-color:black;color:white;padding:0.5em 1.2em;border:none;">⬇️ Download Invoice</button></a>'
    st.markdown(href, unsafe_allow_html=True)
