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
import datetime

# ===== Auto-extraction Function =====
def extract_invoice_details(df_raw):
    """Extract invoice details from Excel sheet using keyword search"""
    
    extracted_data = {}
    
    # Generate PI Number with today's date
    today = datetime.datetime.now()
    pi_date = today.strftime("%d/%m/%Y")
    # Simple counter - in production, you might want a more sophisticated numbering system
    import random
    pi_num = f"SAR/LG/{random.randint(1000, 9999)}"
    extracted_data['pi_number'] = f"{pi_num} Dt. {pi_date}"
    
    # Search through all cells for keywords
    for row_idx, row in df_raw.iterrows():
        for col_idx, cell in enumerate(row):
            if pd.isna(cell):
                continue
            cell_str = str(cell).strip()
            
            # Buyer Name - Row 1, Column A (index 0)
            if row_idx == 0 and col_idx == 0:
                extracted_data['buyer_name'] = cell_str
            
            # Order No - search for "Order No :" and get value 2 cells to the right
            elif "Order No" in cell_str and ":" in cell_str:
                if col_idx + 2 < len(row):
                    order_value = row.iloc[col_idx + 2]
                    if not pd.isna(order_value):
                        extracted_data['order_ref'] = str(order_value).strip()
            
            # Brand Name - search for "Brand" and get value 1 cell to the right
            elif "Brand" in cell_str and cell_str.lower() != "brand name":  # Avoid header matches
                if col_idx + 1 < len(row):
                    brand_value = row.iloc[col_idx + 1]
                    if not pd.isna(brand_value):
                        extracted_data['brand_name'] = str(brand_value).strip()
            
            # Loading Country - search for "Made in Country" and get value 1 cell to the right
            elif "Made in Country" in cell_str:
                if col_idx + 1 < len(row):
                    country_value = row.iloc[col_idx + 1]
                    if not pd.isna(country_value):
                        extracted_data['loading_country'] = str(country_value).strip()
            
            # Port of Loading - search for "Loading Port" and get value 1 cell to the right
            elif "Loading Port" in cell_str:
                if col_idx + 1 < len(row):
                    port_value = row.iloc[col_idx + 1]
                    if not pd.isna(port_value):
                        extracted_data['port_loading'] = str(port_value).strip()
            
            # Agreed Shipment Date - search for "Agreed Ship Date" and get value 2 cells to the right
            elif "Agreed Ship Date" in cell_str:
                if col_idx + 2 < len(row):
                    ship_value = row.iloc[col_idx + 2]
                    if not pd.isna(ship_value):
                        # Handle datetime objects by extracting only the date part
                        if hasattr(ship_value, 'date'):
                            # If it's a datetime object, get just the date
                            extracted_data['shipment_date'] = ship_value.date().strftime('%d/%m/%Y')
                        else:
                            # If it's already a string, clean it up
                            ship_str = str(ship_value).strip()
                            # Remove time portion if present (anything after space)
                            if ' ' in ship_str:
                                ship_str = ship_str.split(' ')[0]
                            extracted_data['shipment_date'] = ship_str
            
            # Description of goods - search for "ORDER OF" and get value 1 cell to the right
            elif "ORDER OF" in cell_str:
                if col_idx + 1 < len(row):
                    goods_value = row.iloc[col_idx + 1]
                    if not pd.isna(goods_value):
                        extracted_data['goods_desc'] = str(goods_value).strip()
    
    return extracted_data

# ===== Hidden Row Detection Function =====
def get_visible_rows_openpyxl(uploaded_file):
    """Get list of visible row indices using openpyxl"""
    try:
        import openpyxl
        from io import BytesIO
        
        # Reset file pointer and read with openpyxl
        uploaded_file.seek(0)
        workbook = openpyxl.load_workbook(BytesIO(uploaded_file.read()))
        worksheet = workbook.active
        
        visible_rows = []
        for row_num in range(1, worksheet.max_row + 1):
            # Check if row is not hidden
            if not worksheet.row_dimensions[row_num].hidden:
                visible_rows.append(row_num - 1)  # Convert to 0-based index for pandas
        
        return visible_rows
    except Exception as e:
        print(f"Could not detect hidden rows with openpyxl: {e}")
        return None

# ===== Preprocessing Function =====
def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    # First, get visible row indices
    visible_row_indices = get_visible_rows_openpyxl(uploaded_file)
    
    # Reset file pointer for pandas
    uploaded_file.seek(0)
    df_raw = pd.read_excel(uploaded_file, header=None)
    
    # Filter to only visible rows if we successfully detected them
    if visible_row_indices is not None:
        # Ensure we don't go beyond dataframe bounds
        max_row_in_df = len(df_raw) - 1
        valid_visible_rows = [r for r in visible_row_indices if r <= max_row_in_df]
        
        if valid_visible_rows:
            print(f"Filtering to {len(valid_visible_rows)} visible rows out of {len(df_raw)} total rows")
            df_raw = df_raw.iloc[valid_visible_rows].reset_index(drop=True)
        else:
            print("No valid visible rows found, using all rows")
    else:
        print("Hidden row detection failed, processing all rows")

    # detect header row
    header_row_idx = None
    stacked_header_idx = None
    for i in range(min(max_rows, len(df_raw))):
        if i >= len(df_raw):
            break
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
    
    # Remove specific unwanted style codes (as backup filter)
    df = df[df["STYLE NO"] != "SA0167A21"]

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

    # Read fabric type from column N, row 5 with bounds checking
    try:
        if len(df_raw) > 4 and len(df_raw.columns) > 13:
            fabric_type_value = df_raw.iloc[4, 13]  # Row 5 (index 4), Column N (index 13)
            if pd.isna(fabric_type_value) or str(fabric_type_value).strip() == "":
                fabric_type_value = "Knitted"  # Default fallback
            else:
                fabric_type_value = str(fabric_type_value).strip()
        else:
            fabric_type_value = "Knitted"  # Default if bounds are exceeded
    except (IndexError, KeyError):
        fabric_type_value = "Knitted"  # Default fallback if cell doesn't exist

    # static extras
    grouped["FABRIC TYPE"] = fabric_type_value
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
                            leftMargin=34.6, rightMargin=34.6)
    elements = []

    styles = getSampleStyleSheet()
    title_style = ParagraphStyle('Title', parent=styles['Normal'], fontSize=12,
                                 alignment=TA_CENTER, fontName='Helvetica-Bold', spaceAfter=6,
                                 borderWidth=1, borderColor=colors.black, borderPadding=(2,6,6,6))
    header_style = ParagraphStyle('Header', parent=styles['Normal'], fontSize=7,
                                  fontName='Helvetica-Bold', alignment=TA_LEFT, 
                                  spaceBefore=0, spaceAfter=0, leading=8)
    normal_style = ParagraphStyle('Normal', parent=styles['Normal'], fontSize=6, alignment=TA_LEFT,
                                  spaceBefore=0, spaceAfter=0, leading=7)

    elements.append(Paragraph("Proforma Invoice", title_style))

    # width setup - adjust product table to align with header sections
    # First calculate the total table width from original product columns to maintain consistency
    original_product_col_widths = [0.8*inch, 1.3*inch, 0.8*inch, 0.7*inch,
                                   1.1*inch, 0.7*inch, 0.5*inch, 0.6*inch, 0.8*inch]
    total_table_width = sum(original_product_col_widths)
    
    # Calculate widths so the line between H.S NO and COMPOSITION aligns with center divider above
    left_section_width = total_table_width/2  # This should align with the center line above
    right_section_width = total_table_width/2
    
    # Distribute left section width among first 4 columns (STYLE NO, ITEM DESC, FABRIC TYPE, H.S NO)
    # Distribute right section width among last 5 columns (COMPOSITION, COUNTRY, QTY, UNIT PRICE, AMOUNT)
    product_col_widths = [
        left_section_width * 0.2,   # STYLE NO (20% of left)
        left_section_width * 0.35,  # ITEM DESCRIPTION (35% of left)  
        left_section_width * 0.25,  # FABRIC TYPE (25% of left)
        left_section_width * 0.2,   # H.S NO (20% of left)
        right_section_width * 0.22, # COMPOSITION (22% of right)
        right_section_width * 0.18, # COUNTRY OF ORIGIN (18% of right) - increased from 15%
        right_section_width * 0.15, # QTY (15% of right)
        right_section_width * 0.2,  # UNIT PRICE (20% of right)
        right_section_width * 0.25  # AMOUNT (25% of right) - reduced from 28%
    ]
    header_col_widths = [total_table_width/2, total_table_width/2]

    # Supplier section
    supplier_data = [
        [Paragraph("<b>Supplier Name:</b><br/><br/>", header_style),  # Added equal spacing to supplier name
         Paragraph(f"<b>No. & date of PI:</b> {form_data['pi_number']}<br/><br/>", header_style)],  # Keep spacing for PI number
        [Paragraph("<b>SAR APPARELS INDIA PVT.LTD.</b><br/><b>Address:</b> 6, Picaso Bithi, Kolkata - 700017<br/><b>Phone:</b> 9817473373<br/><b>Fax:</b> N.A.", ParagraphStyle('SupplierDetail', parent=header_style, leading=12)),
         Paragraph(f"<b>Landmark order Reference:</b> {form_data['order_ref']}<br/><b>Buyer Name:</b> {form_data['buyer_name']}<br/><b>Brand Name:</b> {form_data['brand_name']}", ParagraphStyle('TopAlign', parent=header_style, alignment=TA_LEFT, spaceBefore=0, leading=12))],
    ]
    elements.append(Table(supplier_data, colWidths=header_col_widths,
                          style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                 ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                                 ('LINEBELOW',(1,0),(1,0),1,colors.black),
                                 ('VALIGN',(0,1),(1,1),'TOP'),
                                 ('BOTTOMPADDING',(0,1),(0,1),50),    # Increased bottom padding to 50 points
                                 ('BOTTOMPADDING',(1,1),(1,1),50)]))  # Increased bottom padding to 50 points

    # Consignee section - ULTRA TIGHT SPACING
    # Create compact styles for bank details
    bank_style = ParagraphStyle('BankCompact', parent=normal_style, fontSize=7, fontName='Helvetica', 
                               leading=12, spaceAfter=0, spaceBefore=0, leftIndent=0, rightIndent=0)
    
    consignee_data = [
        [Paragraph("<b>Consignee:</b><br/><br/>", header_style),
         Paragraph(f"<b>Payment Term:</b> {form_data['payment_term']}", normal_style)],
        [Paragraph(f"{form_data['consignee_name']}<br/>{form_data['consignee_address']}<br/>{form_data['consignee_tel']}", 
                   ParagraphStyle('ConsigneeCompact', parent=normal_style, leading=12, spaceAfter=0, spaceBefore=0)),
         Paragraph(f"<br/><br/><b>Bank Details</b><br/><b>BENEFICIARY</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:-&nbsp;{form_data['bank_beneficiary']}<br/><b>ACCOUNT NO</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_account']}<br/><b>BANK'S NAME</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_name']}<br/><b>BANK ADDRESS</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- 2 BRABOURNE ROAD, GOVIND BHAVAN,<br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;GROUND FLOOR, KOLKATA-700001<br/><b>SWIFT CODE</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_swift']}<br/><b>BANK CODE</b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;:- {form_data['bank_code']}", 
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
        # Row 1: Description of goods (left), Currency on right
        [Paragraph(f"<b>Description of goods:</b> {form_data['goods_desc']}", 
                   ParagraphStyle('Goods', parent=normal_style, fontSize=7)), 
         Paragraph("<b>CURRENCY: USD</b>", 
                   ParagraphStyle('Currency', parent=normal_style, 
                                  fontSize=8, alignment=TA_RIGHT, fontName='Helvetica-Bold'))]
    ]
    
    combined_table = Table(combined_data, colWidths=header_col_widths,
                           style=[
                               # Outer border only - NO line between rows
                               ('BOX',(0,0),(-1,-1),1,colors.black),
                               ('LINEBEFORE',(1,0),(1,-1),1,colors.black),
                               ('VALIGN',(0,0),(0,0),'TOP'),      # Description of goods - TOP alignment
                               ('VALIGN',(1,0),(1,0),'BOTTOM')   # Currency - BOTTOM alignment
                           ])
    # Set row height
    combined_table._argH[0] = 50  # Single row height
    elements.append(combined_table)

    # Product Table with additional empty rows
    headers = ["STYLE NO.","ITEM DESCRIPTION","FABRIC TYPE\nKNITTED / WOVEN","H.S NO\n(8digit)",
               "COMPOSITION OF\nMATERIAL","COUNTRY OF\nORIGIN","QTY","UNIT PRICE\nFOB","AMOUNT"]
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

    # Add 5 empty rows for spacing
    for i in range(5):
        table_data.append(["","","","","","","","",""])

    # Function for Indian number formatting
    def indian_format(number):
        """Format number with Indian comma placement (x,xx,xxx pattern)"""
        if number == 0:
            return "0.00"
        
        # Convert to string with 2 decimal places
        num_str = f"{number:.2f}"
        integer_part, decimal_part = num_str.split(".")
        
        # Reverse the integer part for easier processing
        reversed_int = integer_part[::-1]
        
        # Add commas: first comma after 3 digits, then every 2 digits
        formatted = ""
        for i, digit in enumerate(reversed_int):
            if i == 3:  # First comma after 3 digits
                formatted = "," + formatted
            elif i > 3 and (i - 3) % 2 == 0:  # Then every 2 digits
                formatted = "," + formatted
            formatted = digit + formatted
        
        return f"{formatted}.{decimal_part}"

    # TOTAL row with Indian formatting
    table_data.append(
        ["Total","","","","","",f"{total_qty:,}","",f"USD            {indian_format(total_amount)}"]
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
        ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),  # Make TOTAL row bold
        ('WORDWRAP', (0,0), (-1,-1), 'CJK'),  # Enable text wrapping for all cells
    ]))
    elements.append(product_table)

    # Signature block with e-stamp and total in words
    total_words_str = num2words(round(total_amount), to='cardinal', lang='en').upper()
    # Remove commas from the total in words
    total_words_str = total_words_str.replace(",", "")
    total_words_str = f"TOTAL IN WORDS: USD {total_words_str} DOLLARS"

    signature_data = [
        [Paragraph(total_words_str, ParagraphStyle('TotalWords', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=7, alignment=TA_LEFT)), ""],
        [Paragraph("Terms & Conditions (If Any)", ParagraphStyle('TermsCompact', parent=normal_style, spaceBefore=-10)), ""],
        [Image("https://raw.githubusercontent.com/dyas-ai/invoice-generator1/main/Screenshot%202025-09-06%20163303.png", width=2.4*inch, height=1.2*inch), ""],
        ["", ""],  # Empty row for spacing
        [Paragraph("Signed by …………………….(Affix Stamp here)", normal_style),
         Paragraph("&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;for RNA Resources Group Ltd-Landmark (Babyshop)", normal_style)]
    ]
    signature_table = Table(signature_data, colWidths=header_col_widths,
                            style=[('BOX',(0,0),(-1,-1),1,colors.black),
                                   ('VALIGN',(0,-1),(-1,-1),'BOTTOM'),
                                   ('SPAN',(0,0),(-1,0)),
                                   ('BOTTOMPADDING',(0,2),(0,2),15),  # Add bottom padding to stamp row
                                   ('LEFTPADDING',(0,2),(0,2),30),   # Add left padding to move stamp right
                                   ('TOPPADDING',(0,1),(0,1),0),     # Zero top padding for Terms row
                                   ('BOTTOMPADDING',(0,1),(0,1),0),  # Zero bottom padding for Terms row
                                   ('TOPPADDING',(0,2),(0,2),40)])   # Increased top padding to push e-signature down more
    
    # Set specific row heights
    signature_table._argH[1] = 4   # Keep the "Terms & Conditions" row small
    signature_table._argH[2] = 120 # Increase e-signature row height further to restore original spacing
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
        
        # Extract invoice details from Excel
        df_raw = pd.read_excel(uploaded_file, header=None)
        auto_extracted = extract_invoice_details(df_raw)
        
        st.write("### Preview of Processed Data")
        
        # Always update session state with new file data when file is uploaded
        # Store the current file name to detect when a new file is uploaded
        current_file_name = uploaded_file.name
        if 'current_file_name' not in st.session_state or st.session_state.current_file_name != current_file_name:
            st.session_state.edited_df = df.copy()
            st.session_state.current_file_name = current_file_name
        
        # Editable data editor - disable on_change to prevent constant re-runs
        edited_df = st.data_editor(
            st.session_state.edited_df,
            use_container_width=True,
            num_rows="dynamic",
            column_config={
                "QTY": st.column_config.NumberColumn("QTY", format="%d"),
                "UNIT PRICE": st.column_config.NumberColumn("UNIT PRICE", format="%.2f"),
                "AMOUNT": st.column_config.NumberColumn("AMOUNT", format="%.2f")
            },
            disabled=["AMOUNT"],  # Make AMOUNT read-only since it's calculated
            key="data_editor"
        )
        
        # Calculate amounts for the current edited data with proper NaN handling
        working_df = edited_df.copy()
        
        # Clean and handle NaN values before processing
        for col in working_df.columns:
            if col in ["QTY", "UNIT PRICE", "AMOUNT"]:
                # Replace NaN with 0 for numeric columns
                working_df[col] = working_df[col].fillna(0)
            else:
                # Replace NaN with empty string for text columns
                working_df[col] = working_df[col].fillna("")
        
        # Convert numeric columns with proper error handling
        working_df["QTY"] = pd.to_numeric(working_df["QTY"], errors="coerce").fillna(0).astype(int)
        working_df["UNIT PRICE"] = pd.to_numeric(working_df["UNIT PRICE"], errors="coerce").fillna(0.0).astype(float)
        
        # Add text truncation and abbreviations for long content
        def truncate_text(text, max_length=15):
            """Truncate text and add ellipsis if too long"""
            if pd.isna(text) or text == "":
                return ""
            text = str(text).strip()
            if len(text) <= max_length:
                return text
            return text[:max_length-3] + "..."
        
        # Common country abbreviations
        country_abbreviations = {
            "United States of America": "USA",
            "United Kingdom": "UK", 
            "United Arab Emirates": "UAE",
            "Saudi Arabia": "KSA",
            "South Africa": "ZA",
            "New Zealand": "NZ"
        }
        
        # Apply truncation and abbreviations with proper NaN handling
        for idx, row in working_df.iterrows():
            # Handle country names with abbreviations first
            country = str(row.get("COUNTRY OF ORIGIN", "")).strip()
            if country in country_abbreviations:
                working_df.at[idx, "COUNTRY OF ORIGIN"] = country_abbreviations[country]
            else:
                working_df.at[idx, "COUNTRY OF ORIGIN"] = truncate_text(country, 12)
            
            # Truncate other long text fields
            working_df.at[idx, "STYLE NO"] = truncate_text(row.get("STYLE NO", ""), 12)
            working_df.at[idx, "ITEM DESCRIPTION"] = truncate_text(row.get("ITEM DESCRIPTION", ""), 18)
            working_df.at[idx, "FABRIC TYPE"] = truncate_text(row.get("FABRIC TYPE", ""), 12)
            working_df.at[idx, "COMPOSITION"] = truncate_text(row.get("COMPOSITION", ""), 15)
        
        # Calculate amounts after all cleaning is done
        working_df["AMOUNT"] = working_df["QTY"] * working_df["UNIT PRICE"]
        
        # Remove rows where required fields are not filled (but allow zero values)
        # Check if the original data had null/empty values before we filled them with defaults
        rows_to_keep = []
        for idx, row in edited_df.iterrows():
            style_filled = pd.notna(row["STYLE NO"]) and str(row["STYLE NO"]).strip() != ""
            qty_filled = pd.notna(row["QTY"])  # Allow 0 but not NaN/empty
            price_filled = pd.notna(row["UNIT PRICE"])  # Allow 0.0 but not NaN/empty
            
            if style_filled and qty_filled and price_filled:
                rows_to_keep.append(idx)
        
        # Keep only rows that have all required fields filled
        if rows_to_keep:
            working_df = working_df.iloc[rows_to_keep].reset_index(drop=True)
        else:
            # If no valid rows, return empty dataframe with correct columns
            working_df = working_df.iloc[0:0]
        
        # Show summary statistics
        total_qty = working_df["QTY"].sum()
        total_amount = working_df["AMOUNT"].sum()
        st.write(f"**Total Quantity:** {total_qty:,} | **Total Amount:** ${total_amount:,.2f}")

        with st.form("invoice_form"):
            st.subheader("✍️ Enter Invoice Details")
            # Use extracted values as defaults, but allow manual override
            pi_number = st.text_input("PI No. & Date", value=auto_extracted.get('pi_number', 'SAR/LG/XXXX Dt. 10/09/2025'))
            order_ref = st.text_input("Landmark order Reference", value=auto_extracted.get('order_ref', 'CPO/47062/25'))
            buyer_name = st.text_input("Buyer Name", value=auto_extracted.get('buyer_name', 'LANDMARK GROUP'))
            brand_name = st.text_input("Brand Name", value=auto_extracted.get('brand_name', 'Juniors'))
            consignee_name = st.text_input("Consignee Name", value="", placeholder="Enter consignee company name")
            consignee_address = st.text_area("Consignee Address", value="", placeholder="Enter complete consignee address with city, country, postal code")
            consignee_tel = st.text_input("Consignee Tel/Fax", value="", placeholder="Tel: +XXX X XXXXXXX, Fax: +XXX X XXXXXXX")
            payment_term = st.text_input("Payment Term", value="T/T")
            bank_beneficiary = st.text_input("Bank Beneficiary", value="", placeholder="Enter beneficiary company name")
            bank_account = st.text_input("Account No", value="", placeholder="Enter bank account number")
            bank_name = st.text_input("Bank Name", value="", placeholder="Enter bank name")
            bank_address = st.text_area("Bank Address", value="", placeholder="Enter complete bank address with branch, city, country")
            bank_swift = st.text_input("SWIFT", value="", placeholder="Enter SWIFT/BIC code (e.g., KKBKINBBCPC)")
            bank_code = st.text_input("Bank Code", value="", placeholder="Enter bank code/routing number")
            loading_country = st.text_input("Loading Country", value=auto_extracted.get('loading_country', 'India'))
            port_loading = st.text_input("Port of Loading", value=auto_extracted.get('port_loading', 'Mumbai'))
            shipment_date = st.text_input("Agreed Shipment Date", value=auto_extracted.get('shipment_date', '07/02/2025'))
            remarks = st.text_area("Remarks", value="", placeholder="Enter any additional remarks or special instructions (optional)")
            goods_desc = st.text_input("Description of goods", value=auto_extracted.get('goods_desc', 'Value Packs'))
            submitted = st.form_submit_button("Generate PDF")

        if submitted:
            # Update session state only when form is submitted
            st.session_state.edited_df = working_df
            
            form_data = {"pi_number":pi_number,"order_ref":order_ref,"buyer_name":buyer_name,"brand_name":brand_name,
                         "consignee_name":consignee_name,"consignee_address":consignee_address,"consignee_tel":consignee_tel,
                         "payment_term":payment_term,"bank_beneficiary":bank_beneficiary,"bank_account":bank_account,
                         "bank_name":bank_name,"bank_address":bank_address,"bank_swift":bank_swift,"bank_code":bank_code,
                         "loading_country":loading_country,"port_loading":port_loading,"shipment_date":shipment_date,
                         "remarks":remarks,"goods_desc":goods_desc}

            # Use the working dataframe (with calculated amounts) for PDF generation
            pdf_buffer = generate_proforma_invoice(working_df, form_data)
            st.download_button("📥 Download Proforma Invoice PDF", data=pdf_buffer, file_name="proforma_invoice.pdf", mime="application/pdf")

    except Exception as e:
        st.error(f"❌ Error: {e}")
