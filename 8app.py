import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
import io
from datetime import datetime
import num2words

def generate_proforma_invoice(df, pi_number=None, po_reference=None, shipment_date=None):
    """
    Generate proforma invoice PDF from DataFrame matching the exact format
    Returns: PDF bytes
    """
    pdf_buffer = io.BytesIO()
    
    # ===== PDF Setup =====
    styles = getSampleStyleSheet()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, topMargin=0.5*inch, bottomMargin=0.5*inch)
    elements = []
    
    # Custom styles
    title_style = ParagraphStyle(
        'CustomTitle',
        parent=styles['Heading1'],
        fontSize=14,
        spaceAfter=12,
        alignment=1,  # Center
        textColor=colors.black
    )
    
    header_style = ParagraphStyle(
        'HeaderStyle',
        parent=styles['Normal'],
        fontSize=9,
        spaceAfter=3
    )
    
    # Generate PI number if not provided
    if not pi_number:
        pi_number = f"SAR/LG/{datetime.now().strftime('%m%d')}"
    
    # ===== HEADER SECTION =====
    elements.append(Paragraph("Proforma Invoice", title_style))
    elements.append(Spacer(1, 6))
    
    # Create header table with supplier and PI details
    header_data = [
        ["Supplier Name", "No. & date of PI"],
        ["SAR APPARELS INDIA PVT.LTD.", f"{pi_number} Dt. {datetime.now().strftime('%d-%m-%Y')}"],
        ["ADDRESS : 6, Picaso Bithi, KOLKATA - 700017.", f"Landmark order Reference: {po_reference or 'CPO/47062/25'}"],
        ["PHONE : 9874173373", f"Buyer Name: LANDMARK GROUP"],
        ["FAX : N.A.", "Brand Name: Juniors"],
        ["", "Payment Term: T/T"],
        ["Consignee:-", ""],
        ["RNA Resources Group Ltd- Landmark (Babyshop),", "Bank Details (Including Swift/IBAN)"],
        ["P O Box 25030, Dubai, UAE,", ":- SAR APPARELS INDIA PVT.LTD"],
        ["Tel: 00971 4 8095500, Fax: 00971 4 8095555/66", ":- 2112819952"],
        ["", "BANK'S NAME :- KOTAK MAHINDRA BANK LTD"],
        ["", "BANK ADDRESS :- 2 BRABOURNE ROAD, GOVIND BHAVAN, GROUND FLOOR,"],
        ["", "KOLKATA-700001"],
        ["", ":- KKBKINBBCPC"],
        ["", "BANK CODE :- 0323"],
        ["Loading Country: India", "L/C Advicing Bank (If Payment term LC Applicable )"],
        ["Port of loading: Mumbai", ""],
        [f"Agreed Shipment Date: {shipment_date or '07-02-2025'}", ""],
        ["REMARKS if ANY:-", ""],
        ["Description of goods: Value Packs", "CURRENCY: USD"]
    ]
    
    header_table = Table(header_data, colWidths=[4*inch, 4*inch])
    header_table.setStyle(TableStyle([
        ('FONTSIZE', (0,0), (-1,-1), 8),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 2),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2),
    ]))
    
    elements.append(header_table)
    elements.append(Spacer(1, 12))
    
    # ===== MAIN TABLE =====
    table_data = [
        ["STYLE NO.", "ITEM DESCRIPTION", "FABRIC TYPE\nKNITTED /\nWOVEN", "H.S NO\n(8digit)", 
         "COMPOSITION OF\nMATERIAL", "COUNTRY OF\nORIGIN", "QTY", "UNIT PRICE\nFOB", "AMOUNT"]
    ]
    
    total_qty = 0
    total_amount = 0
    
    for _, row in df.iterrows():
        qty = row.get("Qty", 0) or 0
        price = row.get("Unit Price", 0) or 0
        amount = row.get("Amount", 0) or 0
        
        # Handle NaN values
        if pd.isna(qty): qty = 0
        if pd.isna(price): price = 0
        if pd.isna(amount): amount = 0
        
        total_qty += qty
        total_amount += amount
        
        table_data.append([
            str(row.get("StyleID", "") or ""),
            str(row.get("Item Description", "") or ""),
            str(row.get("Fabric Type", "KNITTED") or "KNITTED"),
            str(row.get("HS Code", "") or "").replace(".", ""),  # Remove dots from HS code
            str(row.get("Composition", "") or ""),
            str(row.get("Country of Origin", "India") or "India"),
            f"{qty:,}" if qty > 0 else "",
            f"{price:.2f}" if price > 0 else "",
            f"{amount:.2f}" if amount > 0 else ""
        ])
    
    # Add total row
    table_data.append([
        "", "", "", "", "", "", "", "Total", f"{total_amount:.2f}USD"
    ])
    
    # Create main table
    main_table = Table(table_data, repeatRows=1)
    main_table.setStyle(TableStyle([
        # Header row styling
        ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 7),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        
        # Grid
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        
        # Total row styling
        ('BACKGROUND', (0,-1), (-1,-1), colors.lightgrey),
        ('FONTNAME', (0,-1), (-1,-1), 'Helvetica-Bold'),
        
        # Column alignments
        ('ALIGN', (6,1), (6,-2), 'RIGHT'),  # Qty column
        ('ALIGN', (7,1), (7,-2), 'RIGHT'),  # Price column
        ('ALIGN', (8,1), (8,-1), 'RIGHT'),  # Amount column
        
        # Padding
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
    ]))
    
    elements.append(main_table)
    elements.append(Spacer(1, 12))
    
    # ===== TOTAL IN WORDS =====
    try:
        amount_words = num2words.num2words(int(total_amount), to='currency', currency='USD').upper()
        # Clean up the words format
        amount_words = amount_words.replace(' AND ZERO CENTS', ' DOLLARS')
        if 'CENTS' not in amount_words and 'DOLLARS' not in amount_words:
            amount_words += ' DOLLARS'
    except:
        amount_words = f"TOTAL AMOUNT: ${total_amount:,.2f}"
    
    total_words_style = ParagraphStyle(
        'TotalWords',
        parent=styles['Normal'],
        fontSize=9,
        alignment=1,  # Center
        spaceAfter=12
    )
    
    elements.append(Paragraph(f"TOTAL US DOLLAR {amount_words}", total_words_style))
    elements.append(Spacer(1, 20))
    
    # ===== FOOTER SECTION =====
    footer_data = [
        [f"Total\n{total_qty:,}", "Terms & Conditions (If Any)"],
        ["", ""],
        ["", ""],
        ["Signed by …………………….(Affix Stamp here)", "for RNA Resources Group Ltd-Landmark (Babyshop)"]
    ]
    
    footer_table = Table(footer_data, colWidths=[2*inch, 6*inch])
    footer_table.setStyle(TableStyle([
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('ALIGN', (0,0), (0,0), 'CENTER'),  # Total alignment
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
    ]))
    
    elements.append(footer_table)
    
    # ===== Build PDF =====
    doc.build(elements)
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), total_qty, total_amount

def main():
    st.set_page_config(
        page_title="SAR Apparels - Proforma Invoice Generator",
        page_icon="📄",
        layout="wide"
    )
    
    st.title("📄 SAR Apparels - Proforma Invoice Generator")
    st.markdown("Convert your Excel file to professional Proforma Invoice PDF (Landmark Format)")
    
    # Sidebar for instructions
    with st.sidebar:
        st.header("📋 Instructions")
        st.markdown("""
        1. Upload your Excel file (.xlsx or .xls)
        2. Preview the data to ensure it's correct
        3. Customize invoice details if needed
        4. Click 'Generate PDF' to create your invoice
        5. Download the generated PDF
        
        **Required Excel Columns:**
        - StyleID
        - Item Description
        - Fabric Type
        - HS Code
        - Composition
        - Country of Origin
        - Qty
        - Unit Price
        - Amount
        """)
        
        st.markdown("---")
        st.markdown("**Company:** SAR APPARELS INDIA PVT.LTD.")
        st.markdown("**Client:** Landmark Group (Babyshop)")
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose an Excel file",
        type=['xlsx', 'xls'],
        help="Upload your Excel file containing product details"
    )
    
    if uploaded_file is not None:
        try:
            # Read the Excel file
            df = pd.read_excel(uploaded_file)
            
            # Display file info
            st.success(f"✅ File uploaded successfully! ({len(df)} rows)")
            
            # Show data preview
            st.subheader("📊 Data Preview")
            st.dataframe(df, use_container_width=True)
            
            # Show data summary
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Items", len(df))
            with col2:
                if 'Qty' in df.columns:
                    total_qty = df['Qty'].fillna(0).sum()
                    st.metric("Total Quantity", f"{total_qty:,.0f}")
            with col3:
                if 'Amount' in df.columns:
                    total_amount = df['Amount'].fillna(0).sum()
                    st.metric("Total Amount", f"${total_amount:,.2f}")
            
            # Invoice customization
            st.subheader("🎯 Invoice Settings")
            col1, col2, col3 = st.columns(3)
            
            with col1:
                pi_number = st.text_input(
                    "PI Number",
                    value=f"SAR/LG/{datetime.now().strftime('%m%d')}",
                    help="Proforma Invoice Number"
                )
            
            with col2:
                po_reference = st.text_input(
                    "PO Reference",
                    value="CPO/47062/25",
                    help="Purchase Order Reference"
                )
            
            with col3:
                shipment_date = st.date_input(
                    "Agreed Shipment Date",
                    help="Expected shipment date"
                ).strftime('%d-%m-%Y')
            
            # Generate PDF button
            if st.button("🚀 Generate Proforma Invoice PDF", type="primary", use_container_width=True):
                try:
                    with st.spinner("Generating PDF in Landmark format... Please wait"):
                        pdf_bytes, total_qty, total_amount = generate_proforma_invoice(
                            df, pi_number, po_reference, shipment_date
                        )
                    
                    st.success("✅ PDF generated successfully!")
                    
                    # Display summary
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Total Items", len(df))
                    with col2:
                        st.metric("Total Quantity", f"{total_qty:,.0f}")
                    with col3:
                        st.metric("Total Amount", f"${total_amount:,.2f}")
                    
                    # Download button
                    filename = f"PI_{pi_number.replace('/', '_')}_{datetime.now().strftime('%d-%m-%Y')}"
                    st.download_button(
                        label="📥 Download Proforma Invoice PDF",
                        data=pdf_bytes,
                        file_name=f"{filename}.pdf",
                        mime="application/pdf",
                        type="primary",
                        use_container_width=True
                    )
                    
                except Exception as e:
                    st.error(f"❌ Error generating PDF: {str(e)}")
                    st.write("Please check your Excel file format and try again.")
                    st.exception(e)  # For debugging
            
        except Exception as e:
            st.error(f"❌ Error reading Excel file: {str(e)}")
            st.write("Please ensure your file is a valid Excel format (.xlsx or .xls)")
    
    else:
        st.info("👆 Please upload an Excel file to get started")
        
        # Show sample format based on the reference
        st.subheader("📋 Expected Excel Format (Based on Reference)")
        sample_data = {
            'StyleID': ['SAV001S25', 'SAV002S25', 'SAV003S25'],
            'Item Description': ['S/L Bodysuit 7pk', 'S/L Bodysuit 7pk', 'S/L Bodysuit 7pk'],
            'Fabric Type': ['KNITTED', 'KNITTED', 'KNITTED'],
            'HS Code': ['61112000', '61112000', '61112000'],
            'Composition': ['100% COTTON', '100% COTTON', '100% COTTON'],
            'Country of Origin': ['India', 'India', 'India'],
            'Qty': [4107, 4593, 4593],
            'Unit Price': [6.00, 6.00, 6.00],
            'Amount': [24642.00, 27558.00, 27558.00]
        }
        sample_df = pd.DataFrame(sample_data)
        st.dataframe(sample_df, use_container_width=True)
        
        st.info("💡 The generated PDF will exactly match your reference format including all company details, bank information, and layout structure.")

if __name__ == "__main__":
    main()
