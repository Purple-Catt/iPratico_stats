import pandas as pd
import numpy as np


def analyze_restaurant_orders(file_path):
    """
    Parses an irregular Excel file containing restaurant order histories.
    Each record is separated by two empty lines.
    Handles 'CONTO EMESSO', 'CONTO ALLA ROMANA' (deduplicating identical splits),
    and 'SEPARA CONTO'. Ignores orders without a cover charge ('coperto').

    Parameters:
    file_path (str): Path to the Excel (.xls) file.

    Returns:
    pd.DataFrame: A DataFrame containing the processed and filtered valid orders.
    """
    # 1. Load the .xls file without headers, treating the entire sheet as a raw matrix
    # Note: Ensure 'xlrd' is installed (pip install xlrd) to read old .xls files
    raw_df = pd.read_excel(file_path, header=None)

    records = []
    current_record_rows = []
    empty_row_count = 0

    # 2. Segment the records based on two consecutive empty rows
    for idx, row in raw_df.iterrows():
        is_empty = row.isna().all()

        if is_empty:
            empty_row_count += 1
            # Close the current record block if two consecutive empty rows are found
            if empty_row_count >= 2 and current_record_rows:
                records.append(pd.DataFrame(current_record_rows))
                current_record_rows = []
        else:
            empty_row_count = 0  # Reset counter if the row contains data
            current_record_rows.append(row)

    # Append the last record block if present
    if current_record_rows:
        records.append(pd.DataFrame(current_record_rows))

    print(f"Found {len(records)} raw order blocks in the file.")

    # 3. Filter and normalize records
    valid_orders = []
    romana_signatures = set()  # Set to track and filter out duplicate entries for 'CONTO ALLA ROMANA'
    for rec in records:
        # Clean the block by dropping fully empty rows and columns
        rec = rec.dropna(how='all', axis=0).dropna(how='all', axis=1)
        if rec.empty:
            continue

        # The first cell of each record defines the order type
        first_cell = str(rec.iloc[0, 0]).strip().upper()

        recognized_types = ["DOCUMENTO EMESSO", "CONTO ALLA ROMANA", "SEPARA CONTO"]
        if not any(order_type in first_cell for order_type in recognized_types):
            continue
            continue  # Skip if the header is unrecognized

        # Handle duplicates for 'CONTO ALLA ROMANA'
        if "CONTO ALLA ROMANA" in first_cell:
            # Estraiamo solo le righe che appartengono all'"ORDINE ORIGINALE" per fare il confronto dei duplicati
            original_order_rows = []
            capturing = False
            for _, r in rec.iterrows():
                r_vals = [str(v).strip().upper() for v in r.values if pd.notna(v)]
                if r_vals and r_vals[0] == "ORDINE ORIGINALE":
                    capturing = True
                    continue
                if capturing:
                    # Se incontriamo un'altra sezione o fine blocco, fermiamo la cattura
                    if r_vals and r_vals[0] in ["MODALITA DI PAGAMENTO", "DOCUMENTO EMESSO"]:
                        break
                    original_order_rows.append(r)

            if original_order_rows:
                original_df = pd.DataFrame(original_order_rows)
                record_signature = tuple(original_df.fillna("").astype(str).values.flatten())
            else:
                record_signature = tuple(rec.fillna("").astype(str).values.flatten())

            if record_signature in romana_signatures:
                continue  # Salta il duplicato
            romana_signatures.add(record_signature)

        # 4. Data Extraction Logic

        # Determine starting index for items table
        # For 'CONTO ALLA ROMANA', items are listed under 'ORDINE ORIGINALE'
        start_idx = None
        for i, row in rec.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if "CONTO ALLA ROMANA" in first_cell:
                if row_vals[0] == "ORDINE ORIGINALE":
                    start_idx = i + 1
                    break
            else:
                if row_vals[0] == "OPERATORE - COMANDE":
                    start_idx = i
                    break

        if start_idx is None:
            continue

        # Extract table rows starting from the header until "MODALITA DI PAGAMENTO"
        end_idx = None
        for i, row in rec.loc[start_idx:].iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if row_vals and row_vals[0] == "MODALITA DI PAGAMENTO":
                end_idx = i
                break

        # Slice the dataframe up to the payment method row (exclusive)
        sub_rec = rec.loc[start_idx:end_idx - 1].copy()

        # Locate header row containing column names
        header_row_idx = None
        for i, row in sub_rec.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if "PRODOTTO" in row_vals:
                header_row_idx = i
                break

        if header_row_idx is None:
            continue

        # Map column positions dynamically based on header text
        headers = [str(v).strip() for v in rec.loc[header_row_idx].values]

        prod_col = next((idx for idx, h in enumerate(headers) if "PRODOTTO" in h.strip().upper()), None)
        total_col = next((idx for idx, h in enumerate(headers) if "TOTALE" in h.strip().upper()), None)
        qty_col = next((idx for idx, h in enumerate(headers) if "QUANTITÃ" in h.strip().upper()), None)

        if prod_col is None or total_col is None:
            print("CHE SUCCEDEEEE")
            continue

        total_price = 0.0
        cover_count = 0
        has_cover = False

        # Iterate over item rows
        for i, row in sub_rec.loc[header_row_idx + 1:].iterrows():
            row_vals = [str(v).strip() for v in row.values]
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).upper()

            # Stop if we hit a non-item section
            if "MODALITA DI PAGAMENTO" in row_str or "DOCUMENTO EMESSO" in row_str:
                print("NON DOVREI MAI PRINTARE")
                break

            product_name = str(row_vals[prod_col]).strip()
            total_val_raw = row_vals[total_col]

            # Parse row item total price
            try:
                if pd.notna(total_val_raw):
                    val_str = str(total_val_raw).replace(",", ".").strip()
                    total_price += float(val_str)
            except ValueError:
                print("CHE STRACAZZO SUCCEDE")
                pass

            # Check for coperto
            if "** N. COPERTO" in product_name.upper():
                has_cover = True
                if qty_col is not None:
                    try:
                        qty_str = str(row_vals[qty_col]).strip()
                        cover_count += int(float(qty_str))
                    except ValueError:
                        print("CHE STRACAZZO SUCCEDEEEE PT.2")
                        pass

        # Ignore the order if it contains no coperto
        if not has_cover:
            continue

        valid_orders.append({
            "order_type": first_cell,
            "total_price": round(total_price, 2),
            "cover_count": cover_count,
            "price_per_cover": round(total_price / cover_count, 2) if cover_count > 0 else 0.0,
        })

        # 5. Compile the final dataset
    result_df = pd.DataFrame(valid_orders)

    # --- CREAZIONE DEL SUMMARY DF ---
    if not result_df.empty:
        total_covers = int(result_df["cover_count"].sum())
        total_revenue = result_df["total_price"].sum()

        # Coperto medio calcolato come media pesata (Totale Ricavi / Totale Coperti)
        # oppure come media semplice dei singoli rapporti ("price_per_cover")
        weighted_avg_cover_price = total_revenue / total_covers if total_covers > 0 else 0.0
        simple_avg_cover_price = result_df["price_per_cover"].mean()

        summary_data = [{
            "total_covers": total_covers,
            "total_revenue": round(total_revenue, 2),
            "weighted_avg_cover_price": round(weighted_avg_cover_price, 2),
            "simple_avg_cover_price": round(simple_avg_cover_price, 2),
            "total_orders_analyzed": len(result_df)
        }]

        summary_df = pd.DataFrame(summary_data)

        # Salvataggio o stampa dei risultati
        result_df.to_csv("Risultati_dettaglio.csv", index=False)
        summary_df.to_csv("Summary.csv", index=False)

    else:
        summary_df = pd.DataFrame()
        print("Nessun ordine valido trovato.")

    return result_df, summary_df


if __name__ == "__main__":
    a, b = analyze_restaurant_orders('C:\\Users\\marco\\Desktop\\Code\\shift_planner\\test_ipratico.xlsx')
