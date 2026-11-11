import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import os

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))


def analyze_restaurant_orders(file_path):
    """
    Parses an irregular Excel file containing restaurant order histories.
    Each record is separated by two empty lines.
    Handles 'CONTO EMESSO', 'CONTO ALLA ROMANA' (deduplicating identical splits),
    and 'SEPARA CONTO'. Ignores orders without a cover charge ('coperto').

    Parameters:
        file_path (str): The file path to the Excel (.xlsx) order history file.

    Returns:
        tuple:
            - *result_df* (pd.DataFrame): DataFrame containing the processed valid orders,
              including order type, total price, cover count, and price per cover.
            - *summary_df* (pd.DataFrame): DataFrame containing the overall summary metrics
              (total covers, total revenue, weighted and simple average cover prices,
            and total orders analyzed).
    """
    # 1. Load the .xlsx file
    raw_df = pd.read_excel(file_path, header=None)

    records = []
    current_record_rows = []
    empty_row_count = 0

    # Segment the records based on two consecutive empty rows
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

    # Filter and normalize records
    valid_orders = []
    romana_signatures = set()  # Set to track and filter out duplicate entries for 'CONTO ALLA ROMANA'
    for rec in records:
        rec = rec.dropna(how='all', axis=0).dropna(how='all', axis=1)
        if rec.empty:
            continue

        first_cell = str(rec.iloc[0, 0]).strip().upper()

        recognized_types = ["DOCUMENTO EMESSO", "CONTO ALLA ROMANA", "SEPARA CONTO"]
        if not any(order_type in first_cell for order_type in recognized_types):
            continue

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
                continue

            romana_signatures.add(record_signature)

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
            raise NameError("Your file is inconsistent. Check the data")

        # Extract table rows starting from the header until "MODALITA DI PAGAMENTO"
        end_idx = None
        for i, row in rec.loc[start_idx:].iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]
            if row_vals and row_vals[0] == "MODALITA DI PAGAMENTO":
                end_idx = i
                break

        sub_rec = rec.loc[start_idx:end_idx - 1].copy()

        # Locate header row containing column names
        header_row_idx = None
        for i, row in sub_rec.iterrows():
            row_vals = [str(v).strip().upper() for v in row.values if pd.notna(v)]

            if "PRODOTTO" in row_vals:
                header_row_idx = i
                break

        if header_row_idx is None:
            raise KeyError("Key 'PRODOTTO' is missing, check your .xlsx file!")

        # Map column positions dynamically based on header text
        headers = [str(v).strip() for v in rec.loc[header_row_idx].values]

        prod_col = next((idx for idx, h in enumerate(headers) if "PRODOTTO" in h.strip().upper()), None)
        total_col = next((idx for idx, h in enumerate(headers) if "TOTALE" in h.strip().upper()), None)
        qty_col = next((idx for idx, h in enumerate(headers) if "QUANTITÃ" in h.strip().upper()), None)

        if prod_col is None or total_col is None:
            raise KeyError("Price or Total column(s) missing")

        total_price = 0.0
        cover_count = 0
        has_cover = False

        # Iterate over item rows
        for i, row in sub_rec.loc[header_row_idx + 1:].iterrows():
            row_vals = [str(v).strip() for v in row.values]
            row_str = " ".join([str(v) for v in row.values if pd.notna(v)]).upper()

            if "MODALITA DI PAGAMENTO" in row_str or "DOCUMENTO EMESSO" in row_str:
                break

            product_name = str(row_vals[prod_col]).strip()
            total_val_raw = row_vals[total_col]

            # Parse row item total price
            try:
                if pd.notna(total_val_raw):
                    val_str = str(total_val_raw).replace(",", ".").strip()
                    total_price += float(val_str)
            except ValueError:
                pass

            # Check for cover charge
            if "** N. COPERTO" in product_name.upper():
                has_cover = True
                if qty_col is not None:
                    try:
                        qty_str = str(row_vals[qty_col]).strip()
                        cover_count += int(float(qty_str))
                    except ValueError:
                        pass

        # Ignore the order if it contains no cover charge
        if not has_cover:
            continue

        valid_orders.append({
            "order_type": first_cell,
            "total_price": round(total_price, 2),
            "cover_count": cover_count,
            "price_per_cover": round(total_price / cover_count, 2) if cover_count > 0 else 0.0,
        })

        # Compile the final dataset
    result_df = pd.DataFrame(valid_orders)

    if not result_df.empty:
        total_covers = int(result_df["cover_count"].sum())
        total_revenue = result_df["total_price"].sum()

        # Coperto medio calcolato come media pesata (Totale Ricavi / Totale Coperti)
        # oppure come media semplice dei singoli rapporti ("price_per_cover")
        weighted_avg_cover_price = total_revenue / total_covers if total_covers > 0 else 0.0
        simple_avg_cover_price = result_df["price_per_cover"].mean()

        summary_data = [{
            "Totale coperti": total_covers,
            "Totale incasso": round(total_revenue, 2),
            "Prezzo medio (ponderato) per coperto": round(weighted_avg_cover_price, 2),
            "Prezzo medio (semplice) per coperto": round(simple_avg_cover_price, 2),
            "Totale scontrini analizzati": len(result_df)
        }]

        summary_df = pd.DataFrame(summary_data)

        # Salvataggio nella stessa cartella da cui viene eseguito l'exe
        base_dir = os.path.dirname(os.path.abspath(__file__)) if '__file__' in locals() else os.getcwd()
        res_path = os.path.join(base_dir, "RISULTATI_TEST.csv")
        sum_path = os.path.join(base_dir, "SUMMARY_TEST.csv")

        result_df.to_csv(res_path, index=False)
        summary_df.to_csv(sum_path, index=False)

        return True, res_path
    else:
        return False, None


if __name__ == "__main__":
    # Nasconde la finestra principale di tkinter, mostra solo il selettore file
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Seleziona il file Excel delle comande",
        filetypes=[("File Excel", "*.xlsx"), ("Tutti i file", "*.*")]
    )

    if file_path:
        try:
            success, path = analyze_restaurant_orders(file_path)
            if success:
                messagebox.showinfo("Successo",
                                    f"Elaborazione completata con successo!\n\nFile "
                                    f"salvati in:\n{os.path.dirname(path)}")
            else:
                messagebox.showwarning("Attenzione", "Nessun ordine valido con coperto trovato nel file.")
        except Exception as e:
            messagebox.showerror("Errore", f"Si è verificato un errore durante l'elaborazione:\n{str(e)}")
