def preprocess_excel_flexible_auto(uploaded_file, max_rows=20):
    df_raw = pd.read_excel(uploaded_file, header=None)

    header_row_idx = None
    stacked_header_idx = None

    # Step 1: Detect header row (row containing "Style")
    for i in range(min(max_rows, len(df_raw))):
        row = df_raw.iloc[i].astype(str)
        if row.str.contains("Style", case=False, na=False).any():
            header_row_idx = i
            stacked_header_idx = i - 1
            break

    if header_row_idx is None:
        raise ValueError("Could not detect header row with 'Style' column!")

    # Step 2: Combine stacked headers safely (convert everything to string)
    if stacked_header_idx >= 0:
        headers = df_raw.iloc[stacked_header_idx].astype(str).fillna('') + ' ' + df_raw.iloc[header_row_idx].astype(str).fillna('')
    else:
        headers = df_raw.iloc[header_row_idx].astype(str).fillna('')

    headers = headers.str.strip()

    # Step 3: Define possible column names for mapping
    col_map = {
        "STYLE NO": ["Style", "Style No", "Item Style"],
        "ITEM DESCRIPTION": ["Descreption", "Description", "Item Description", "Item Desc"],
        "COMPOSITION": ["Composition", "Fabric Composition"],
        "UNIT PRICE": ["Fob$", "USD Fob$", "Fob USD", "Fob $"],
        "QTY": ["Total Qty", "Quantity", "Qty"],
        "AMOUNT": ["Total Value", "Amount", "Value"]
    }

    # Step 4: Map columns automatically
    df_columns = {}
    for target_col, variants in col_map.items():
        for var in variants:
            matched_cols = [c for c in headers if var.lower() in str(c).lower()]
            if matched_cols:
                df_columns[target_col] = matched_cols[0]
                break
        if target_col not in df_columns:
            df_columns[target_col] = None  # Column not found

    # Step 5: Select data rows (rows after header)
    df = df_raw.iloc[header_row_idx + 1:].copy()
    df.columns = headers
    df = df.reset_index(drop=True)

    # Step 6: Rename columns to standard names
    rename_dict = {v: k for k, v in df_columns.items() if v is not None}
    df = df.rename(columns=rename_dict)

    # ✅ Step 6.1: Drop rows without STYLE NO (removes totals/notes)
    if "STYLE NO" in df.columns:
        df["STYLE NO"] = df["STYLE NO"].astype(str)
        df = df[~df["STYLE NO"].str.contains("total|grand|remarks|note", case=False, na=False)]
        df = df[df["STYLE NO"].notna()]
        df = df[df["STYLE NO"].str.strip() != ""]

    # Step 7: Convert numeric fields
    df["QTY"] = pd.to_numeric(df["QTY"], errors="coerce").fillna(0).astype(int)
    df["UNIT PRICE"] = pd.to_numeric(df["UNIT PRICE"], errors="coerce").fillna(0.0)

    # ✅ Step 7.1: Drop junk rows where Qty & Amount are zero but no Style
    df = df[~((df["QTY"] == 0) & (df["UNIT PRICE"] == 0) & (df["STYLE NO"] == ""))]

    # Step 8: Aggregate per unique style
    grouped = (
        df.groupby(["STYLE NO", "ITEM DESCRIPTION", "COMPOSITION", "UNIT PRICE"], dropna=False)
        .agg({"QTY": "sum"})
        .reset_index()
    )

    # Step 9: Compute Amount = QTY × UNIT PRICE
    grouped["AMOUNT"] = grouped["QTY"] * grouped["UNIT PRICE"]

    # Step 10: Add static PDF columns
    grouped["FABRIC TYPE"] = ""
    grouped["HS CODE"] = ""
    grouped["COUNTRY OF ORIGIN"] = "India"

    # Step 11: Reorder columns for PDF
    grouped = grouped[
        ["STYLE NO", "ITEM DESCRIPTION", "FABRIC TYPE", "HS CODE",
         "COMPOSITION", "COUNTRY OF ORIGIN", "QTY", "UNIT PRICE", "AMOUNT"]
    ]

    return grouped
