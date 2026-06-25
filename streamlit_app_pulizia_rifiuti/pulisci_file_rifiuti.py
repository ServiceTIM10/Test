from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Any

from openpyxl import load_workbook


# =========================
# CONFIGURAZIONE BASE
# =========================
HEADER_ROW = 1
COL_NR_DOC = "Nr. Doc."
COL_MC = "MC"
COL_KG = "Kg"

# Formati Excel: Excel li visualizza secondo la lingua/locale dell'utente.
# In Excel italiano, #,##0 diventa 1.234 e #,##0.00 diventa 1.234,56.
FORMAT_NR_DOC = "#,##0"
FORMAT_QTA = "#,##0.00"


# =========================
# FUNZIONI DI SUPPORTO
# =========================
def unwrap_excel_text_formula(value: str) -> str:
    """
    Gestisce eventuali valori salvati come formule testuali del tipo ="...".
    Esempio: ="5.640" -> 5.640
    """
    s = value.strip()
    if s.startswith('="') and s.endswith('"'):
        return s[2:-1]
    if s.startswith("='") and s.endswith("'"):
        return s[2:-1]
    return s


def normalize_nr_doc(value: Any) -> int | None:
    """
    Normalizza la colonna Nr. Doc. in numero intero.

    Regola principale per il problema del file:
    - se il valore è numerico non intero, ad esempio 5.64 o 5.645,
      significa che il separatore delle migliaia è stato interpretato come separatore decimale;
      quindi 5.64 -> 5.640 e 5.645 -> 5.645, cioè si moltiplica per 1000.

    Esempi:
    - 5.64    -> 5640
    - 5.645   -> 5645
    - 1.23    -> 1230
    - 1229    -> 1229
    - '5,64'  -> 5640
    - '5.640' -> 5640
    """
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        raise ValueError(f"Valore Nr. Doc. non valido: {value!r}")

    if isinstance(value, int):
        return value

    if isinstance(value, float):
        if value.is_integer():
            return int(value)
        return int(round(value * 1000))

    s = unwrap_excel_text_formula(str(value)).strip()
    s = s.replace(" ", "").replace("\u00a0", "")
    s = s.strip("'")

    if not s:
        return None

    # Solo cifre: già intero.
    if re.fullmatch(r"\d+", s):
        return int(s)

    # Formato italiano corretto con punto migliaia: 1.230, 5.640, 12.345
    if re.fullmatch(r"\d{1,3}(\.\d{3})+", s):
        return int(s.replace(".", ""))

    # Caso troncato come stringa: 5,64 -> 5.640 / 1,23 -> 1.230
    if re.fullmatch(r"\d+,\d+", s):
        return int(round(float(s.replace(",", ".")) * 1000))

    # Caso troncato con punto decimale come stringa: 5.64 -> 5.640
    if re.fullmatch(r"\d+\.\d+", s):
        return int(round(float(s) * 1000))

    # Formati misti, es. 1.234,00: rimuovo migliaia e poi considero la parte intera.
    if re.fullmatch(r"\d{1,3}(\.\d{3})+,\d+", s):
        numeric = float(s.replace(".", "").replace(",", "."))
        if not numeric.is_integer():
            raise ValueError(f"Nr. Doc. dovrebbe essere intero, trovato: {value!r}")
        return int(numeric)

    raise ValueError(f"Formato Nr. Doc. non riconosciuto: {value!r}")


def parse_italian_or_standard_number(value: Any) -> float | None:
    """
    Converte MC e Kg in numero float gestendo formati standard, italiani
    e casi sporchi/malformati.

    Esempi gestiti:
    - 1234.56        -> 1234.56
    - '1234.56'      -> 1234.56
    - '1.234,56'     -> 1234.56
    - '1234,56'      -> 1234.56
    - '1,234.56'     -> 1234.56
    - '1.125.00'     -> 1125.00
    - '1.125.000'    -> 1125000.00
    - '1.125.000.00' -> 1125000.00
    - '' / None      -> None
    """
    if value is None or value == "":
        return None

    if isinstance(value, bool):
        raise ValueError(f"Valore numerico non valido: {value!r}")

    if isinstance(value, (int, float)):
        return float(value)

    s_originale = str(value)

    s = unwrap_excel_text_formula(s_originale).strip()
    s = s.replace(" ", "").replace("\u00a0", "")
    s = s.strip("'")
    s = s.replace("€", "")

    if not s:
        return None

    # Rimuove eventuali caratteri non visibili o separatori strani
    s = s.replace("\t", "").replace("\n", "").replace("\r", "")

    if "," in s and "." in s:
        # Il separatore decimale è quello più a destra.
        if s.rfind(",") > s.rfind("."):
            # Italiano: 1.234,56 -> 1234.56
            s = s.replace(".", "").replace(",", ".")
        else:
            # Inglese/internazionale: 1,234.56 -> 1234.56
            s = s.replace(",", "")

    elif "," in s:
        # Italiano senza migliaia: 1234,56 -> 1234.56
        s = s.replace(",", ".")

    elif "." in s:
        if s.count(".") > 1:
            parts = s.split(".")

            # Caso sporco tipo:
            # 1.125.00 -> 1125.00
            # 1.125.000.00 -> 1125000.00
            if len(parts[-1]) in (1, 2) and all(part.isdigit() for part in parts):
                integer_part = "".join(parts[:-1])
                decimal_part = parts[-1]
                s = integer_part + "." + decimal_part

            # Caso con soli separatori migliaia:
            # 1.125.000 -> 1125000
            elif all(part.isdigit() for part in parts) and all(len(part) == 3 for part in parts[1:]):
                s = "".join(parts)

            else:
                raise ValueError(
                    f"Formato numerico con punti non riconosciuto: valore originale={value!r}, valore pulito={s!r}"
                )

    # Ultima protezione: se per qualsiasi motivo arriva ancora un valore tipo 1.125.00,
    # lo correggiamo immediatamente prima della conversione finale.
    if s.count(".") > 1 and "," not in s:
        parts = s.split(".")

        if len(parts[-1]) in (1, 2) and all(part.isdigit() for part in parts):
            s = "".join(parts[:-1]) + "." + parts[-1]
        elif all(part.isdigit() for part in parts) and all(len(part) == 3 for part in parts[1:]):
            s = "".join(parts)

    try:
        return float(s)
    except ValueError as e:
        raise ValueError(
            f"Impossibile convertire il valore numerico: valore originale={value!r}, valore dopo pulizia={s!r}"
        ) from e

def get_header_map(ws) -> dict[str, int]:
    """Restituisce {nome_colonna: indice_colonna_excel_1_based}."""
    headers = {}
    for cell in ws[HEADER_ROW]:
        if cell.value is not None:
            headers[str(cell.value).strip()] = cell.column
    return headers


# =========================
# FUNZIONE PRINCIPALE
# =========================
def clean_excel_file(input_path: str | Path, output_path: str | Path, sheet_name: str | None = None) -> None:
    input_path = Path(input_path)
    output_path = Path(output_path)

    wb = load_workbook(input_path)
    ws = wb[sheet_name] if sheet_name else wb.active

    headers = get_header_map(ws)
    required = [COL_NR_DOC, COL_MC, COL_KG]
    missing = [col for col in required if col not in headers]
    if missing:
        raise ValueError(f"Colonne mancanti nel file: {missing}")

    col_doc = headers[COL_NR_DOC]
    col_mc = headers[COL_MC]
    col_kg = headers[COL_KG]

    data_rows = []
    corrected_doc_count = 0

    # Lettura efficiente dei dati: iter_rows è molto più rapido di ws.cell(...) riga per riga.
    for row in ws.iter_rows(
        min_row=HEADER_ROW + 1,
        max_row=ws.max_row,
        max_col=ws.max_column,
        values_only=True,
    ):
        values = list(row)

        old_doc = values[col_doc - 1]
        new_doc = normalize_nr_doc(old_doc)

        if old_doc != new_doc:
            corrected_doc_count += 1

        values[col_doc - 1] = new_doc
        values[col_mc - 1] = parse_italian_or_standard_number(values[col_mc - 1])
        values[col_kg - 1] = parse_italian_or_standard_number(values[col_kg - 1])

        data_rows.append(values)

    # Ordinamento crescente dell'intera tabella sulla base di Nr. Doc.
    # Le righe con Nr. Doc. vuoto, se presenti, vengono messe in fondo.
    data_rows.sort(key=lambda row: (row[col_doc - 1] is None, row[col_doc - 1] or 0))

    # Riscrittura dei dati ordinati mantenendo intestazioni e struttura del foglio.
    for out_row_idx, values in enumerate(data_rows, start=HEADER_ROW + 1):
        for col_idx, value in enumerate(values, start=1):
            ws.cell(out_row_idx, col_idx).value = value

    # Applicazione dei formati numerici.
    for row_idx in range(HEADER_ROW + 1, ws.max_row + 1):
        ws.cell(row_idx, col_doc).number_format = FORMAT_NR_DOC
        ws.cell(row_idx, col_mc).number_format = FORMAT_QTA
        ws.cell(row_idx, col_kg).number_format = FORMAT_QTA

    # Imposto larghezze minime utili.
    ws.column_dimensions[ws.cell(HEADER_ROW, col_doc).column_letter].width = 12
    ws.column_dimensions[ws.cell(HEADER_ROW, col_mc).column_letter].width = 12
    ws.column_dimensions[ws.cell(HEADER_ROW, col_kg).column_letter].width = 12

    wb.save(output_path)

    print(f"File pulito salvato in: {output_path}")
    print(f"Righe elaborate: {len(data_rows)}")
    print(f"Valori Nr. Doc. convertiti/normalizzati: {corrected_doc_count}")


# =========================
# ESECUZIONE DA TERMINALE
# =========================
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Pulisce file Excel rifiuti: Nr. Doc., MC e Kg.")
    parser.add_argument("input", help="Percorso del file Excel da pulire")
    parser.add_argument("output", help="Percorso del file Excel pulito da generare")
    parser.add_argument("--sheet", default=None, help="Nome del foglio da pulire. Se omesso usa il primo foglio.")
    args = parser.parse_args()

    clean_excel_file(args.input, args.output, args.sheet)


#gg
